import json
import smtplib
import ssl
from datetime import datetime, date, timedelta
import pandas as pd
import yfinance as yf
import pandas_market_calendars as mcal
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl

def load_config():
    with open('C:\\TradingBot\\config.json') as f:
        return json.load(f)

def is_market_open_today():
    nyse = mcal.get_calendar('NYSE')
    today = datetime.now().strftime('%Y-%m-%d')
    schedule = nyse.schedule(start_date=today, end_date=today)
    return not schedule.empty

def get_next_friday():
    today = datetime.now()
    days_until_friday = (4 - today.weekday()) % 7
    if days_until_friday == 0:
        days_until_friday = 7
    next_friday = today + pd.Timedelta(days=days_until_friday)
    return next_friday.strftime('%Y-%m-%d')

def get_monthly_expiries():
    today = date.today()
    expiries = []
    for months_ahead in [1, 2]:
        year = today.year + (today.month + months_ahead - 1) // 12
        month = (today.month + months_ahead - 1) % 12 + 1
        first = date(year, month, 1)
        day = first
        fridays = []
        while day.month == month:
            if day.weekday() == 4:
                fridays.append(day)
            day += timedelta(days=1)
        if len(fridays) >= 3:
            third_friday = fridays[2]
            dte = (third_friday - today).days
            expiries.append((third_friday.strftime('%Y-%m-%d'), dte))
    return expiries

def get_stock_data(ticker):
    try:
        stock = yf.Ticker(ticker)
        info = stock.info
        hist = stock.history(period='1y')
        if hist.empty:
            return None
        current = info.get('currentPrice') or info.get('regularMarketPrice') or hist['Close'].iloc[-1]
        pre_market = info.get('preMarketPrice') or current
        week_high = hist['Close'].max()
        prev_close = hist['Close'].iloc[-2] if len(hist) > 1 else current
        pre_market_change_pct = (pre_market - prev_close) / prev_close
        high_proximity_pct = (week_high - current) / week_high
        return {
            'ticker': ticker,
            'current': round(current, 2),
            'pre_market': round(pre_market, 2),
            'prev_close': round(prev_close, 2),
            'week_high': round(week_high, 2),
            'pre_market_change_pct': round(pre_market_change_pct * 100, 2),
            'high_proximity_pct': round(high_proximity_pct * 100, 2)
        }
    except Exception as ex:
        print('  Could not fetch ' + ticker + ': ' + str(ex))
        return None

def get_earnings_tickers(watchlist):
    earnings_soon = []
    for ticker in watchlist:
        try:
            stock = yf.Ticker(ticker)
            cal = stock.calendar
            if cal is not None and not cal.empty:
                earn_date = pd.Timestamp(cal.iloc[0, 0])
                days_away = (earn_date - pd.Timestamp.now()).days
                if 0 <= days_away <= 5:
                    earnings_soon.append((ticker, days_away))
        except:
            pass
    return earnings_soon

def check_sympathy_drop(ticker, peer_group, all_data, threshold):
    peers_down = 0
    peers_checked = 0
    for peer in peer_group:
        if peer == ticker or peer not in all_data or all_data[peer] is None:
            continue
        peers_checked += 1
        if all_data[peer]['pre_market_change_pct'] <= -(threshold * 100):
            peers_down += 1
    return peers_checked >= 2 and peers_down >= 2

def find_put_candidates(cfg, all_data, earnings_tickers, open_positions):
    rules = cfg['rules']
    exclusions = cfg['exclusions']
    tier1 = cfg['tiers']['tier1']
    tier2 = cfg['tiers']['tier2']
    peer_groups = cfg['peer_groups']
    candidates = []
    earnings_tickers_list = [e[0] for e in earnings_tickers]
    for ticker, data in all_data.items():
        if data is None:
            continue
        if ticker in exclusions:
            continue
        if ticker in earnings_tickers_list:
            continue
        if ticker in open_positions:
            continue
        if ticker not in tier1 and ticker not in tier2:
            continue
        if data['high_proximity_pct'] > rules['high_proximity_pct'] * 100:
            continue
        drop_pct = data['pre_market_change_pct']
        if drop_pct >= -(rules['sympathy_drop_pct'] * 100):
            continue
        ticker_group = None
        for group_name, members in peer_groups.items():
            if ticker in members:
                ticker_group = group_name
                break
        if ticker_group is None:
            continue
        is_sympathy = check_sympathy_drop(
            ticker, peer_groups[ticker_group], all_data, rules['sympathy_drop_pct']
        )
        if not is_sympathy:
            continue
        tier = 'Tier 1' if ticker in tier1 else 'Tier 2'
        notional_min = rules['tier1_notional_min'] if ticker in tier1 else rules['tier2_notional_min']
        notional_max = rules['tier1_notional_max'] if ticker in tier1 else rules['tier2_notional_max']
        avg_notional = (notional_min + notional_max) / 2
        strike_low = round(data['current'] * (1 - rules['put_strike_max_otm']), 2)
        strike_high = round(data['current'] * (1 - rules['put_strike_min_otm']), 2)
        strike_mid = (strike_low + strike_high) / 2
        contracts = max(1, round(avg_notional / (data['current'] * 100)))
        est_premium_per_contract = round(strike_mid * rules['min_weekly_premium_pct'], 2)
        est_total_premium = round(est_premium_per_contract * contracts * 100, 2)
        candidates.append({
            'ticker': ticker,
            'tier': tier,
            'group': ticker_group,
            'current': data['current'],
            'pre_market': data['pre_market'],
            'drop_pct': drop_pct,
            'week_high': data['week_high'],
            'proximity_pct': data['high_proximity_pct'],
            'strike_low': strike_low,
            'strike_high': strike_high,
            'contracts': contracts,
            'notional_min': notional_min,
            'notional_max': notional_max,
            'est_premium': est_total_premium,
            'est_per_contract': est_premium_per_contract
        })
    return candidates

def find_longer_dated_candidates(cfg, all_data, earnings_tickers, open_positions):
    rules = cfg['rules']
    exclusions = cfg['exclusions']
    peer_groups = cfg['peer_groups']
    longer_tier1 = ['MSFT', 'AMZN', 'GOOGL', 'META', 'NVDA', 'AVGO']
    candidates = []
    earnings_tickers_list = [e[0] for e in earnings_tickers]
    expiries = get_monthly_expiries()
    for ticker, data in all_data.items():
        if data is None:
            continue
        if ticker in exclusions:
            continue
        if ticker in earnings_tickers_list:
            continue
        if ticker in open_positions:
            continue
        if ticker not in longer_tier1:
            continue
        if data['high_proximity_pct'] > rules['high_proximity_pct'] * 100:
            continue
        drop_pct = data['pre_market_change_pct']
        if drop_pct >= -(rules['sympathy_drop_pct'] * 100):
            continue
        ticker_group = None
        for group_name, members in peer_groups.items():
            if ticker in members:
                ticker_group = group_name
                break
        if ticker_group is None:
            continue
        is_sympathy = check_sympathy_drop(
            ticker, peer_groups[ticker_group], all_data, rules['sympathy_drop_pct']
        )
        if not is_sympathy:
            continue
        current = data['current']
        exp_results = []
        for exp_date, dte in expiries:
            if dte < 25 or dte > 65:
                continue
            weeks = round(dte / 7, 1)
            delta_proxy = 0.12 + (dte / 71) * 0.08
            strike = round(current * (1 - delta_proxy), 2)
            notional = round(strike * 100, 2)
            time_factor = (dte / 365) ** 0.5
            est_prem = round(current * 0.20 * time_factor * 0.20, 2)
            roi = round(est_prem / strike * 100, 2)
            weekly_roi = round(roi / weeks, 2)
            exp_results.append({
                'exp_date': exp_date,
                'dte': dte,
                'weeks': weeks,
                'strike': strike,
                'notional': notional,
                'est_prem': est_prem,
                'roi': roi,
                'weekly_roi': weekly_roi
            })
        if not exp_results:
            continue
        candidates.append({
            'ticker': ticker,
            'group': ticker_group,
            'current': current,
            'drop_pct': drop_pct,
            'week_high': data['week_high'],
            'proximity_pct': data['high_proximity_pct'],
            'expiries': exp_results
        })
    return candidates

def load_positions():
    try:
        wb = openpyxl.load_workbook('C:\\TradingBot\\positions.xlsx')
        puts_ws = wb['OpenPuts']
        assigned_ws = wb['AssignedPositions']
        open_puts = []
        assigned = []
        puts_headers = [cell.value for cell in puts_ws[1]]
        for row in puts_ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            put = {}
            for i, h in enumerate(puts_headers):
                put[h] = row[i]
            open_puts.append(put)
        assigned_headers = [cell.value for cell in assigned_ws[1]]
        for row in assigned_ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            pos = {}
            for i, h in enumerate(assigned_headers):
                pos[h] = row[i]
            assigned.append(pos)
        return open_puts, assigned
    except Exception as ex:
        print('Warning: Could not read positions.xlsx: ' + str(ex))
        return [], []

def check_stops(assigned, all_data):
    alerts = []
    for pos in assigned:
        ticker = pos['Ticker']
        cost_basis = float(pos['CostBasis'])
        highest = float(pos['HighestPriceSeen']) if pos['HighestPriceSeen'] else cost_basis
        if ticker not in all_data or all_data[ticker] is None:
            continue
        current = all_data[ticker]['current']
        static_stop = round(cost_basis * 0.95, 2)
        trailing_active = current >= cost_basis * 1.10 or highest >= cost_basis * 1.10
        trailing_stop = round(highest * 0.95, 2)
        stop_price = trailing_stop if trailing_active else static_stop
        stop_type = 'TRAILING' if trailing_active else 'STATIC'
        pnl_pct = round((current - cost_basis) / cost_basis * 100, 2)
        has_covered_call = pos['CoveredCallStrike'] is not None
        status = 'OK'
        if current <= stop_price:
            if has_covered_call:
                status = 'STOP HIT - BUY BACK CALL FIRST'
            else:
                status = 'STOP HIT - SELL SHARES'
        elif current <= stop_price * 1.03:
            status = 'APPROACHING STOP - WATCH CLOSELY'
        alerts.append({
            'ticker': ticker,
            'current': current,
            'cost_basis': cost_basis,
            'pnl_pct': pnl_pct,
            'stop_price': stop_price,
            'stop_type': stop_type,
            'trailing_active': trailing_active,
            'highest': highest,
            'has_covered_call': has_covered_call,
            'covered_call_strike': pos['CoveredCallStrike'],
            'covered_call_expiry': pos['CoveredCallExpiry'],
            'status': status
        })
    return alerts

def get_call_recommendations(assigned, all_data, rules):
    recommendations = []
    for pos in assigned:
        ticker = pos['Ticker']
        cost_basis = float(pos['CostBasis'])
        shares = int(pos['Shares'])
        highest = float(pos['HighestPriceSeen']) if pos['HighestPriceSeen'] else cost_basis
        has_covered_call = pos['CoveredCallStrike'] is not None
        if has_covered_call:
            continue
        if ticker not in all_data or all_data[ticker] is None:
            continue
        current = all_data[ticker]['current']
        static_stop = round(cost_basis * 0.95, 2)
        trailing_active = highest >= cost_basis * 1.10
        trailing_stop = round(highest * 0.95, 2)
        stop_price = trailing_stop if trailing_active else static_stop
        pnl_pct = round((current - cost_basis) / cost_basis * 100, 2)
        if current <= stop_price * 1.03:
            continue
        if current >= cost_basis:
            mode = 'NORMAL'
            call_strike = round(current * (1 + rules['call_strike_min_otm']), 2)
            call_strike_high = round(current * (1 + rules['call_strike_max_otm']), 2)
            est_premium = round(current * rules['min_call_premium_pct'], 2)
        elif current >= cost_basis * (1 - rules['recovery_mode_threshold']):
            mode = 'RECOVERY'
            call_strike = round(current, 2)
            call_strike_high = round(current * 1.01, 2)
            est_premium = round(current * rules['min_call_premium_pct'] * 1.5, 2)
        else:
            continue
        est_total_premium = round(est_premium * shares, 2)
        recommendations.append({
            'ticker': ticker,
            'current': current,
            'cost_basis': cost_basis,
            'pnl_pct': pnl_pct,
            'shares': shares,
            'mode': mode,
            'call_strike': call_strike,
            'call_strike_high': call_strike_high,
            'est_premium': est_premium,
            'est_total_premium': est_total_premium,
            'stop_price': stop_price,
            'stop_type': 'TRAILING' if trailing_active else 'STATIC'
        })
    return recommendations

def build_report(cfg, candidates, earnings_tickers, all_data, watchlist_flags):
    now = datetime.now().strftime('%A %B %d, %Y %I:%M %p')
    L = []
    L.append('=' * 60)
    L.append('TRADING BOT MORNING REPORT')
    L.append(now)
    L.append('=' * 60)
    L.append('')
    L.append('SECTION A - PORTFOLIO SNAPSHOT')
    L.append('-' * 40)
    total = cfg['portfolio']['total_value']
    reserve = cfg['portfolio']['dry_powder_reserve']
    L.append('Total portfolio value:    $' + format(total, ','))
    L.append('Dry powder reserve:       $' + format(reserve, ','))
    L.append('Available for puts:       $' + format(total - reserve, ','))
    L.append('Note: Update portfolio value in config.json as it changes.')
    L.append('')
    L.append('SECTION B - WEEKLY PUT CANDIDATES TODAY')
    L.append('-' * 40)
    if not candidates:
        L.append('No qualifying weekly put candidates found today.')
        L.append('Reasons: no sympathy drops, earnings soon, or too far from 52W high.')
    else:
        for c in candidates:
            actual_notional = c['contracts'] * c['current'] * 100
            L.append('Ticker:           ' + c['ticker'] + ' (' + c['tier'] + ' | ' + c['group'] + ')')
            L.append('Pre-market price: $' + str(c['pre_market']) + ' (' + str(c['drop_pct']) + '% vs prev close)')
            L.append('Current price:    $' + str(c['current']))
            L.append('52W high:         $' + str(c['week_high']) + ' (' + str(c['proximity_pct']) + '% below high)')
            L.append('Strike range:     $' + str(c['strike_low']) + ' to $' + str(c['strike_high']))
            L.append('Contracts:        ' + str(c['contracts']))
            L.append('Actual notional:  $' + format(int(actual_notional), ','))
            L.append('Est. min premium: $' + str(c['est_premium']) + ' total ($' + str(c['est_per_contract']) + '/contract)')
            L.append('*** Verify actual premium in ATP before placing order ***')
            L.append('')
        L.append('ATP ORDER TICKETS')
        L.append('-' * 40)
        next_friday = get_next_friday()
        for c in candidates:
            L.append('Action:      SELL TO OPEN PUT')
            L.append('Ticker:      ' + c['ticker'])
            L.append('Expiry:      ' + next_friday + ' (verify in ATP)')
            L.append('Strike:      $' + str(c['strike_low']) + ' to $' + str(c['strike_high']) + ' - choose closest standard strike')
            L.append('Contracts:   ' + str(c['contracts']))
            L.append('Order type:  LIMIT at market bid (verify in ATP)')
            L.append('')
    L.append('SECTION E - WATCHLIST FLAGS')
    L.append('-' * 40)
    if earnings_tickers:
        L.append('Earnings within 5 days - DO NOT write puts on these:')
        for ticker, days in earnings_tickers:
            L.append('  ' + ticker + ': ' + str(days) + ' day(s) away')
    else:
        L.append('No earnings alerts today.')
    L.append('')
    if watchlist_flags:
        L.append('Large pre-market moves (>5%):')
        for flag in watchlist_flags:
            direction = 'UP' if flag['change'] > 0 else 'DOWN'
            L.append('  ' + flag['ticker'] + ': ' + str(flag['change']) + '% ' + direction)
    else:
        L.append('No large pre-market moves today.')
    L.append('')
    return '\n'.join(L)

def build_longer_dated_section(candidates):
    L = []
    L.append('SECTION B2 - LONGER DATED PUT CANDIDATES (30-60 DTE)')
    L.append('-' * 40)
    L.append('Tier 1 names only: MSFT AMZN GOOGL META NVDA AVGO')
    L.append('Delta target: ~0.20 | Strikes: 12-20% OTM')
    L.append('NOTE: Premium estimates are approximations only.')
    L.append('Verify actual premiums and deltas in ATP before trading.')
    L.append('')
    if not candidates:
        L.append('No qualifying longer-dated candidates today.')
        L.append('Same sympathy drop and proximity filters apply.')
    else:
        for c in candidates:
            L.append('Ticker:        ' + c['ticker'] + ' (' + c['group'] + ')')
            L.append('Current price: $' + str(c['current']))
            L.append('Pre-mkt move:  ' + str(c['drop_pct']) + '% vs prev close')
            L.append('52W high:      $' + str(c['week_high']) + ' (' + str(c['proximity_pct']) + '% below high)')
            L.append('')
            for e in c['expiries']:
                L.append('  Expiry:      ' + e['exp_date'] + ' (' + str(e['dte']) + ' DTE / ' + str(e['weeks']) + ' weeks)')
                L.append('  Strike:      $' + str(e['strike']))
                L.append('  Notional:    $' + format(int(e['notional']), ','))
                L.append('  Est premium: $' + str(e['est_prem']) + ' per contract')
                L.append('  Total ROI:   ' + str(e['roi']) + '%')
                L.append('  Weekly ROI:  ' + str(e['weekly_roi']) + '% (vs 1% weekly target)')
                L.append('')
            if c['expiries']:
                best = c['expiries'][0]
                L.append('  ATP ORDER TICKET')
                L.append('  Action:      SELL TO OPEN PUT')
                L.append('  Ticker:      ' + c['ticker'])
                L.append('  Expiry:      ' + best['exp_date'] + ' (verify in ATP)')
                L.append('  Strike:      $' + str(best['strike']) + ' (verify delta ~0.20 in ATP)')
                L.append('  Contracts:   1')
                L.append('  Order type:  LIMIT at market bid (verify in ATP)')
            L.append('')
            L.append('-' * 40)
    return '\n'.join(L)

def build_sections_cd(open_puts, assigned, all_data, rules):
    L = []
    L.append('SECTION C - COVERED CALL OPPORTUNITIES')
    L.append('-' * 40)
    if not assigned:
        L.append('No assigned positions on file.')
    else:
        call_recs = get_call_recommendations(assigned, all_data, rules)
        if not call_recs:
            L.append('No covered call opportunities today.')
        else:
            next_friday = get_next_friday()
            for r in call_recs:
                L.append('Ticker:       ' + r['ticker'])
                L.append('Current:      $' + str(r['current']))
                L.append('Cost basis:   $' + str(r['cost_basis']))
                L.append('PnL:          ' + str(r['pnl_pct']) + '%')
                L.append('Mode:         ' + r['mode'])
                L.append('Call strike:  $' + str(r['call_strike']) + ' to $' + str(r['call_strike_high']))
                L.append('Est premium:  $' + str(r['est_premium']) + ' per share')
                L.append('Total prem:   $' + str(r['est_total_premium']))
                L.append('Stop price:   $' + str(r['stop_price']))
                L.append('Stop type:    ' + r['stop_type'])
                L.append('*** Verify premium in ATP before placing order ***')
                L.append('')
            L.append('ATP ORDER TICKETS - COVERED CALLS')
            L.append('-' * 40)
            for r in call_recs:
                L.append('Action:      SELL TO OPEN CALL')
                L.append('Ticker:      ' + r['ticker'])
                L.append('Expiry:      ' + next_friday + ' (verify in ATP)')
                L.append('Strike:      $' + str(r['call_strike']) + ' to $' + str(r['call_strike_high']))
                L.append('Contracts:   ' + str(int(r['shares'] / 100)))
                L.append('Order type:  LIMIT at market bid (verify in ATP)')
                L.append('')
    L.append('SECTION D - STOP ALERTS AND POSITION STATUS')
    L.append('-' * 40)
    if not assigned:
        L.append('No assigned positions on file.')
    else:
        stop_alerts = check_stops(assigned, all_data)
        has_alerts = False
        for a in stop_alerts:
            if a['status'] != 'OK':
                has_alerts = True
                L.append('*** ' + a['status'] + ' ***')
                L.append('Ticker:     ' + a['ticker'])
                L.append('Current:    $' + str(a['current']))
                L.append('Cost basis: $' + str(a['cost_basis']))
                L.append('Stop price: $' + str(a['stop_price']))
                L.append('Stop type:  ' + a['stop_type'])
                L.append('PnL:        ' + str(a['pnl_pct']) + '%')
                if a['has_covered_call']:
                    L.append('WARNING: BUY BACK CALL BEFORE SELLING SHARES')
                L.append('')
        if not has_alerts:
            L.append('No stop alerts. All positions within normal range.')
            L.append('')
            for a in stop_alerts:
                L.append('Ticker: ' + a['ticker'] + ' | Current: $' + str(a['current']) + ' | Stop: $' + str(a['stop_price']) + ' | PnL: ' + str(a['pnl_pct']) + '%')
    L.append('')
    L.append('OPEN PUTS SUMMARY')
    L.append('-' * 40)
    if not open_puts:
        L.append('No open put positions on file.')
    else:
        for p in open_puts:
            L.append('Ticker: ' + str(p['Ticker']) + ' | Strike: $' + str(p['Strike']) + ' | Expiry: ' + str(p['Expiry']) + ' | Contracts: ' + str(p['Contracts']) + ' | Premium: $' + str(p['PremiumCollected']))
    return '\n'.join(L)

def get_performance_summary():
    try:
        wb = openpyxl.load_workbook('C:\\TradingBot\\positions.xlsx')
        ws = wb['ClosedTrades']
        headers = [cell.value for cell in ws[1]]
        trades = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            trade = {}
            for i, h in enumerate(headers):
                trade[h] = row[i]
            trades.append(trade)
        if not trades:
            return None
        total = len(trades)
        expired = len([t for t in trades if t['Outcome'] == 'EXPIRED'])
        assigned = len([t for t in trades if t['Outcome'] == 'ASSIGNED'])
        called_away = len([t for t in trades if t['Outcome'] == 'CALLED_AWAY'])
        stop_losses = len([t for t in trades if t['Outcome'] == 'STOP_LOSS'])
        wins = expired + called_away
        win_rate = round(wins / total * 100, 1) if total > 0 else 0
        returns = [float(t['ReturnPct']) for t in trades if t['ReturnPct'] and float(t['ReturnPct']) > 0]
        avg_return = round(sum(returns) / len(returns) * 100, 2) if returns else 0
        best = round(max(returns) * 100, 2) if returns else 0
        worst = round(min(returns) * 100, 2) if returns else 0
        total_income = sum([float(t['TotalIncome']) for t in trades if t['TotalIncome']])
        return {
            'total': total,
            'expired': expired,
            'assigned': assigned,
            'called_away': called_away,
            'stop_losses': stop_losses,
            'win_rate': win_rate,
            'avg_return': avg_return,
            'best': best,
            'worst': worst,
            'total_income': round(total_income, 2)
        }
    except Exception as ex:
        print('Warning: Could not read performance data: ' + str(ex))
        return None

def build_performance_section():
    L = []
    L.append('SECTION F - PERFORMANCE SUMMARY')
    L.append('-' * 40)
    perf = get_performance_summary()
    if perf is None:
        L.append('No closed trades on file yet.')
    else:
        L.append('Total trades:           ' + str(perf['total']))
        L.append('Puts expired worthless: ' + str(perf['expired']))
        L.append('Puts assigned:          ' + str(perf['assigned']))
        L.append('Covered calls closed:   ' + str(perf['called_away']))
        L.append('Stop losses:            ' + str(perf['stop_losses']))
        L.append('Win rate:               ' + str(perf['win_rate']) + '%')
        L.append('Avg return per trade:   ' + str(perf['avg_return']) + '%')
        L.append('Best trade:             ' + str(perf['best']) + '%')
        L.append('Worst trade:            ' + str(perf['worst']) + '%')
        L.append('Total income to date:   $' + format(perf['total_income'], ','))
    L.append('')
    L.append('=' * 60)
    L.append('IMPORTANT: All recommendations are estimates only.')
    L.append('Verify all premiums and strikes in Fidelity ATP')
    L.append('before placing any order. This is not financial advice.')
    L.append('=' * 60)
    return '\n'.join(L)

def send_email(cfg, subject, body):
    e = cfg['email']
    msg = MIMEMultipart()
    msg['From'] = e['sender']
    msg['To'] = e['recipient']
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx) as server:
        server.login(e['sender'], e['app_password'])
        server.sendmail(e['sender'], e['recipient'], msg.as_string())

def main():
    print('Loading config...')
    cfg = load_config()
    rules = cfg['rules']
    print('Checking market calendar...')
    if not is_market_open_today():
        print('Market is closed today. No report generated.')
        return
    print('Loading positions...')
    open_puts, assigned = load_positions()
    open_position_tickers = [p['Ticker'] for p in open_puts] + [a['Ticker'] for a in assigned]
    watchlist = cfg['watchlist']
    print('Fetching data for ' + str(len(watchlist)) + ' stocks...')
    all_data = {}
    for ticker in watchlist:
        print('  Fetching ' + ticker + '...')
        all_data[ticker] = get_stock_data(ticker)
    print('Checking earnings calendar...')
    earnings_tickers = get_earnings_tickers(watchlist)
    print('Identifying watchlist flags...')
    watchlist_flags = []
    for ticker, data in all_data.items():
        if data and abs(data['pre_market_change_pct']) >= 5.0:
            watchlist_flags.append({'ticker': ticker, 'change': data['pre_market_change_pct']})
    print('Finding put candidates...')
    candidates = find_put_candidates(cfg, all_data, earnings_tickers, open_position_tickers)
    longer_candidates = find_longer_dated_candidates(cfg, all_data, earnings_tickers, open_position_tickers)
    print('Building report...')
    report = build_report(cfg, candidates, earnings_tickers, all_data, watchlist_flags)
    longer_section = build_longer_dated_section(longer_candidates)
    sections_cd = build_sections_cd(open_puts, assigned, all_data, rules)
    perf_section = build_performance_section()
    full_report = report + '\n' + longer_section + '\n' + sections_cd + '\n' + perf_section
    print('')
    print(full_report)
    print('')
    print('Sending email...')
    today = datetime.now().strftime('%Y-%m-%d')
    send_email(cfg, 'Trading Bot Morning Report - ' + today, full_report)
    print('Report emailed successfully.')

if __name__ == '__main__':
    main()

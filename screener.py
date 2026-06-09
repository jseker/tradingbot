import json
import smtplib
import ssl
from datetime import datetime, date, timedelta
import pandas as pd
import numpy as np
import yfinance as yf
import pandas_market_calendars as mcal
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import openpyxl
import urllib.request

# ── MACRO CALENDAR ───────────────────────────────────────────────────
# Static calendar of recurring major economic events
# Format: (month, day_type, day_value, event_name, impact)
# day_type: 'fixed' = specific date, 'weekday' = nth weekday of month
# This covers the most market-moving events for tech/growth stocks

def get_macro_events(start_date, end_date):
    events = []
    current = start_date.replace(day=1)
    while current <= end_date + timedelta(days=31):
        year = current.year
        month = current.month

        # ── CPI (Consumer Price Index) ───────────────────────────────
        # Usually released 2nd or 3rd Wednesday, ~12th-15th of month
        cpi_date = get_nth_weekday(year, month, 2, 2)  # 2nd Wednesday
        if start_date <= cpi_date <= end_date:
            events.append({
                'date': cpi_date,
                'name': 'CPI (Consumer Price Index)',
                'impact': 'HIGH',
                'note': 'Major inflation indicator — can move tech stocks 2-4%'
            })

        # ── PPI (Producer Price Index) ───────────────────────────────
        # Usually day after CPI
        ppi_date = cpi_date + timedelta(days=1)
        if start_date <= ppi_date <= end_date:
            events.append({
                'date': ppi_date,
                'name': 'PPI (Producer Price Index)',
                'impact': 'MEDIUM',
                'note': 'Leading indicator for future CPI'
            })

        # ── Non-Farm Payrolls ─────────────────────────────────────────
        # First Friday of month
        nfp_date = get_nth_weekday(year, month, 1, 4)  # 1st Friday
        if start_date <= nfp_date <= end_date:
            events.append({
                'date': nfp_date,
                'name': 'Non-Farm Payrolls (Jobs Report)',
                'impact': 'HIGH',
                'note': 'Strong jobs = rate fears, weak jobs = recession fears'
            })

        # ── FOMC meetings ─────────────────────────────────────────────
        # Approximately every 6-7 weeks — hardcode 2026 dates
        fomc_2026 = [
            date(2026, 1, 28), date(2026, 3, 18), date(2026, 5, 6),
            date(2026, 6, 17), date(2026, 7, 29), date(2026, 9, 16),
            date(2026, 11, 4), date(2026, 12, 16)
        ]
        for fd in fomc_2026:
            if start_date <= fd <= end_date:
                events.append({
                    'date': fd,
                    'name': 'FOMC Rate Decision',
                    'impact': 'HIGH',
                    'note': 'Fed rate decision — major market mover'
                })

        # ── PCE Inflation ─────────────────────────────────────────────
        # Last Friday of month
        pce_date = get_last_weekday(year, month, 4)  # Last Friday
        if start_date <= pce_date <= end_date:
            events.append({
                'date': pce_date,
                'name': 'PCE Inflation',
                'impact': 'HIGH',
                'note': "Fed's preferred inflation measure"
            })

        # ── Retail Sales ──────────────────────────────────────────────
        # Usually mid-month, around 15th-17th
        retail_date = get_nth_weekday(year, month, 2, 1)  # 2nd Tuesday
        if start_date <= retail_date <= end_date:
            events.append({
                'date': retail_date,
                'name': 'Retail Sales',
                'impact': 'MEDIUM',
                'note': 'Consumer spending indicator'
            })

        # ── ISM Manufacturing ─────────────────────────────────────────
        # First business day of month
        ism_date = get_first_business_day(year, month)
        if start_date <= ism_date <= end_date:
            events.append({
                'date': ism_date,
                'name': 'ISM Manufacturing PMI',
                'impact': 'MEDIUM',
                'note': 'Economic activity indicator'
            })

        # ── Consumer Confidence ───────────────────────────────────────
        # Last Tuesday of month
        conf_date = get_last_weekday(year, month, 1)  # Last Tuesday
        if start_date <= conf_date <= end_date:
            events.append({
                'date': conf_date,
                'name': 'Consumer Confidence',
                'impact': 'LOW',
                'note': 'Consumer sentiment survey'
            })

        # ── Weekly Jobless Claims ─────────────────────────────────────
        # Every Thursday — only flag if within 7 days to avoid clutter
        thursday = get_nth_weekday(year, month, 1, 3)
        while thursday.month == month:
            if start_date <= thursday <= end_date:
                days_away = (thursday - start_date).days
                if days_away <= 10:
                    events.append({
                        'date': thursday,
                        'name': 'Weekly Jobless Claims',
                        'impact': 'LOW',
                        'note': 'Weekly labor market health check'
                    })
            thursday += timedelta(days=7)

        # Move to next month
        if month == 12:
            current = current.replace(year=year+1, month=1)
        else:
            current = current.replace(month=month+1)

    # Deduplicate and sort
    seen = set()
    unique_events = []
    for e in sorted(events, key=lambda x: x['date']):
        key = (e['date'], e['name'])
        if key not in seen:
            seen.add(key)
            unique_events.append(e)

    return unique_events

def get_nth_weekday(year, month, n, weekday):
    # weekday: 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri
    first = date(year, month, 1)
    first_weekday = first.weekday()
    days_until = (weekday - first_weekday) % 7
    first_occurrence = first + timedelta(days=days_until)
    return first_occurrence + timedelta(weeks=n-1)

def get_last_weekday(year, month, weekday):
    if month == 12:
        last = date(year+1, 1, 1) - timedelta(days=1)
    else:
        last = date(year, month+1, 1) - timedelta(days=1)
    days_back = (last.weekday() - weekday) % 7
    return last - timedelta(days=days_back)

def get_first_business_day(year, month):
    d = date(year, month, 1)
    while d.weekday() >= 5:
        d += timedelta(days=1)
    return d

def build_macro_section(open_puts):
    L = []
    L.append('SECTION H - MACRO ECONOMIC CALENDAR')
    L.append('-' * 40)

    today = date.today()

    # Find furthest expiry across all open puts
    max_expiry = today + timedelta(days=30)
    for p in open_puts:
        try:
            exp = p.get('Expiry')
            if exp:
                if hasattr(exp, 'date'):
                    exp = exp.date()
                elif isinstance(exp, datetime):
                    exp = exp.date()
                if exp > max_expiry:
                    max_expiry = exp
        except:
            pass

    days_out = (max_expiry - today).days
    L.append('Scanning macro events from today through ' + max_expiry.strftime('%B %d, %Y'))
    L.append('(' + str(days_out) + ' days — covers full duration of longest open position)')
    L.append('')

    events = get_macro_events(today, max_expiry)

    if not events:
        L.append('No major economic events found in this period.')
        return '\n'.join(L)

    # Group by impact
    high_impact = [e for e in events if e['impact'] == 'HIGH']
    medium_impact = [e for e in events if e['impact'] == 'MEDIUM']
    low_impact = [e for e in events if e['impact'] == 'LOW']

    if high_impact:
        L.append('*** HIGH IMPACT EVENTS ***')
        for e in high_impact:
            days_away = (e['date'] - today).days
            day_label = 'TODAY' if days_away == 0 else ('TOMORROW' if days_away == 1 else str(days_away) + ' days away')
            L.append(e['date'].strftime('%a %b %d') + ' (' + day_label + ') — ' + e['name'])
            L.append('  Impact: ' + e['note'])
            # Flag which open positions are affected
            affected = []
            for p in open_puts:
                try:
                    exp = p.get('Expiry')
                    if exp:
                        if hasattr(exp, 'date'):
                            exp_date = exp.date()
                        elif isinstance(exp, datetime):
                            exp_date = exp.date()
                        else:
                            exp_date = exp
                        if exp_date >= e['date']:
                            affected.append(p['Ticker'] + ' ' + str(p.get('Strike', '')) + ' exp ' + exp_date.strftime('%b %d'))
                except:
                    pass
            if affected:
                L.append('  Affects open positions: ' + ', '.join(affected))
            L.append('')

    if medium_impact:
        L.append('MEDIUM IMPACT EVENTS')
        for e in medium_impact:
            days_away = (e['date'] - today).days
            day_label = str(days_away) + ' days away'
            L.append(e['date'].strftime('%a %b %d') + ' (' + day_label + ') — ' + e['name'])
        L.append('')

    if low_impact:
        L.append('LOW IMPACT EVENTS (within 10 days)')
        for e in low_impact:
            days_away = (e['date'] - today).days
            L.append(e['date'].strftime('%a %b %d') + ' (' + str(days_away) + ' days) — ' + e['name'])
        L.append('')

    # Overall risk warning
    next_7_days = [e for e in high_impact if (e['date'] - today).days <= 7]
    if next_7_days:
        L.append('*** CAUTION: ' + str(len(next_7_days)) + ' HIGH IMPACT event(s) within 7 days ***')
        L.append('Consider reducing new position size or waiting until after announcements.')
    else:
        L.append('No high impact events in next 7 days — clear to write new positions.')

    return '\n'.join(L)

# ── TECHNICAL ANALYSIS / MEAN REVERSION ─────────────────────────────

def calculate_technicals(ticker):
    try:
        stock = yf.Ticker(ticker)
        hist = stock.history(period='3mo')
        if len(hist) < 20:
            return None

        closes = hist['Close']
        volumes = hist['Volume']

        # RSI (14-period)
        delta = closes.diff()
        gain = delta.where(delta > 0, 0)
        loss = -delta.where(delta < 0, 0)
        avg_gain = gain.rolling(14).mean()
        avg_loss = loss.rolling(14).mean()
        rs = avg_gain / avg_loss
        rsi = 100 - (100 / (1 + rs))
        current_rsi = round(float(rsi.iloc[-1]), 1)

        # Bollinger Bands (20-period, 2 std)
        sma20 = closes.rolling(20).mean()
        std20 = closes.rolling(20).std()
        upper_band = sma20 + (2 * std20)
        lower_band = sma20 - (2 * std20)
        current_price = float(closes.iloc[-1])
        current_sma20 = float(sma20.iloc[-1])
        current_lower = float(lower_band.iloc[-1])
        current_upper = float(upper_band.iloc[-1])

        # Distance from moving averages
        sma50 = closes.rolling(50).mean()
        current_sma50 = float(sma50.iloc[-1]) if len(closes) >= 50 else None

        pct_from_sma20 = round((current_price - current_sma20) / current_sma20 * 100, 2)
        pct_from_sma50 = round((current_price - current_sma50) / current_sma50 * 100, 2) if current_sma50 else None

        # Bollinger Band position (0 = lower band, 1 = upper band)
        bb_position = round((current_price - current_lower) / (current_upper - current_lower), 2) if (current_upper - current_lower) > 0 else 0.5

        # Volume analysis — today vs 20-day average
        avg_volume = float(volumes.rolling(20).mean().iloc[-1])
        current_volume = float(volumes.iloc[-1])
        volume_ratio = round(current_volume / avg_volume, 2) if avg_volume > 0 else 1.0

        # 1-day and 5-day price change
        change_1d = round((current_price - float(closes.iloc[-2])) / float(closes.iloc[-2]) * 100, 2) if len(closes) >= 2 else 0
        change_5d = round((current_price - float(closes.iloc[-6])) / float(closes.iloc[-6]) * 100, 2) if len(closes) >= 6 else 0

        return {
            'ticker': ticker,
            'current': round(current_price, 2),
            'rsi': current_rsi,
            'sma20': round(current_sma20, 2),
            'sma50': round(current_sma50, 2) if current_sma50 else None,
            'lower_band': round(current_lower, 2),
            'upper_band': round(current_upper, 2),
            'bb_position': bb_position,
            'pct_from_sma20': pct_from_sma20,
            'pct_from_sma50': pct_from_sma50,
            'volume_ratio': volume_ratio,
            'change_1d': change_1d,
            'change_5d': change_5d
        }
    except Exception as ex:
        print('  Could not calculate technicals for ' + ticker + ': ' + str(ex))
        return None

def assess_mean_reversion(tech, cfg):
    if tech is None:
        return None

    score = 0
    signals = []
    cautions = []

    # RSI signal (most important)
    if tech['rsi'] < 25:
        score += 3
        signals.append('RSI ' + str(tech['rsi']) + ' — extremely oversold')
    elif tech['rsi'] < 30:
        score += 2
        signals.append('RSI ' + str(tech['rsi']) + ' — oversold')
    elif tech['rsi'] < 35:
        score += 1
        signals.append('RSI ' + str(tech['rsi']) + ' — approaching oversold')
    elif tech['rsi'] > 70:
        cautions.append('RSI ' + str(tech['rsi']) + ' — overbought, not a put writing opportunity')
        return None

    # Bollinger Band signal
    if tech['bb_position'] < 0.05:
        score += 3
        signals.append('Price at/below lower Bollinger Band — statistically extreme')
    elif tech['bb_position'] < 0.15:
        score += 2
        signals.append('Price near lower Bollinger Band (BB position: ' + str(tech['bb_position']) + ')')
    elif tech['bb_position'] < 0.25:
        score += 1
        signals.append('Price below lower quarter of Bollinger Bands')

    # Distance from SMA20
    if tech['pct_from_sma20'] < -10:
        score += 2
        signals.append(str(abs(tech['pct_from_sma20'])) + '% below 20-day moving average')
    elif tech['pct_from_sma20'] < -5:
        score += 1
        signals.append(str(abs(tech['pct_from_sma20'])) + '% below 20-day moving average')

    # Volume analysis
    if tech['volume_ratio'] < 0.7:
        score += 1
        signals.append('Low volume drop (' + str(tech['volume_ratio']) + 'x avg) — suggests technical not fundamental')
    elif tech['volume_ratio'] > 2.0:
        cautions.append('High volume drop (' + str(tech['volume_ratio']) + 'x avg) — may signal genuine selling pressure')
        score -= 1

    # 5-day change
    if tech['change_5d'] < -10:
        score += 1
        signals.append(str(tech['change_5d']) + '% drop over 5 days — extended move, bounce likely')

    # Minimum score threshold
    if score < 2:
        return None

    # Conviction level
    if score >= 5:
        conviction = 'HIGH CONVICTION'
    elif score >= 3:
        conviction = 'MODERATE'
    else:
        conviction = 'WATCH'

    # Suggested strikes
    current = tech['current']
    weekly_strike = round(current * 0.95, 2)  # 5% OTM for weekly
    monthly_strike = round(current * 0.88, 2)  # 12% OTM for monthly

    return {
        'ticker': tech['ticker'],
        'conviction': conviction,
        'score': score,
        'signals': signals,
        'cautions': cautions,
        'tech': tech,
        'weekly_strike': weekly_strike,
        'monthly_strike': monthly_strike
    }

def find_mean_reversion_candidates(cfg, all_data, earnings_tickers, open_positions):
    tier1 = cfg['tiers']['tier1']
    tier2 = cfg['tiers']['tier2']
    exclusions = cfg['exclusions']
    earnings_list = [e[0] for e in earnings_tickers]
    candidates = []

    for ticker in tier1 + tier2:
        if ticker in exclusions:
            continue
        if ticker in earnings_list:
            continue
        if ticker in open_positions:
            continue
        if all_data.get(ticker) is None:
            continue

        # Check proximity to 52W high
        data = all_data[ticker]
        if data['high_proximity_pct'] > cfg['rules']['high_proximity_pct'] * 100:
            continue

        # Must have dropped at least 2% recently (not a flat stock)
        if data['pre_market_change_pct'] > 0 and data['high_proximity_pct'] < 3:
            continue

        print('  Calculating technicals for ' + ticker + '...')
        tech = calculate_technicals(ticker)
        if tech is None:
            continue

        # Must have dropped without peer confirmation (isolating from sympathy drops)
        # Check if peers are also down — if yes it is sympathy not mean reversion
        peer_groups = cfg['peer_groups']
        ticker_group = None
        for group_name, members in peer_groups.items():
            if ticker in members:
                ticker_group = group_name
                break

        if ticker_group:
            peers_also_down = 0
            for peer in peer_groups[ticker_group]:
                if peer == ticker or peer not in all_data or all_data[peer] is None:
                    continue
                if all_data[peer]['pre_market_change_pct'] < -2.0:
                    peers_also_down += 1
            # If 2+ peers also down, it is sympathy — skip (handled by Section B)
            if peers_also_down >= 2:
                continue

        result = assess_mean_reversion(tech, cfg)
        if result is not None:
            candidates.append(result)

    # Sort by conviction score
    candidates.sort(key=lambda x: x['score'], reverse=True)
    return candidates

def build_mean_reversion_section(candidates, earnings_tickers):
    L = []
    L.append('SECTION G - MEAN REVERSION CANDIDATES')
    L.append('-' * 40)
    L.append('Stocks that have dropped for no apparent fundamental reason.')
    L.append('Technical analysis suggests potential bounce / mean reversion.')
    L.append('These are ISOLATED drops — not confirmed by peers (see Section B for sympathy drops).')
    L.append('')

    if not candidates:
        L.append('No mean reversion candidates today.')
        L.append('Either no significant isolated drops, or RSI/Bollinger signals not triggered.')
        return '\n'.join(L)

    next_friday = get_next_friday()

    for c in candidates:
        tech = c['tech']
        L.append('*** ' + c['ticker'] + ' — ' + c['conviction'] + ' (Score: ' + str(c['score']) + '/8) ***')
        L.append('Current price:      $' + str(tech['current']))
        L.append('RSI (14):           ' + str(tech['rsi']) + ' ' + ('OVERSOLD' if tech['rsi'] < 30 else ('Near oversold' if tech['rsi'] < 35 else '')))
        L.append('20-day SMA:         $' + str(tech['sma20']) + ' (' + str(tech['pct_from_sma20']) + '% from price)')
        if tech['sma50']:
            L.append('50-day SMA:         $' + str(tech['sma50']) + ' (' + str(tech['pct_from_sma50']) + '% from price)')
        L.append('Lower Bollinger:    $' + str(tech['lower_band']))
        L.append('Upper Bollinger:    $' + str(tech['upper_band']))
        L.append('BB Position:        ' + str(tech['bb_position']) + ' (0=lower band, 1=upper band)')
        L.append('Volume ratio:       ' + str(tech['volume_ratio']) + 'x average')
        L.append('1-day change:       ' + str(tech['change_1d']) + '%')
        L.append('5-day change:       ' + str(tech['change_5d']) + '%')
        L.append('')
        L.append('Technical signals:')
        for s in c['signals']:
            L.append('  + ' + s)
        if c['cautions']:
            L.append('Cautions:')
            for ca in c['cautions']:
                L.append('  ! ' + ca)
        L.append('')
        L.append('Suggested put strikes:')
        L.append('  Weekly (' + next_friday + '): $' + str(c['weekly_strike']) + ' (~5% OTM)')
        L.append('  Monthly (next expiry):   $' + str(c['monthly_strike']) + ' (~12% OTM)')
        L.append('*** Verify delta, premium and IV in ATP before trading ***')
        L.append('')
        L.append('-' * 40)

    return '\n'.join(L)

# ── EXISTING FUNCTIONS ────────────────────────────────────────────────

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

def clean_header(h):
    if h is None:
        return None
    return str(h).replace('\n', '').replace(' ', '')

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
        puts_headers = [clean_header(cell.value) for cell in puts_ws[3]]
        for row in puts_ws.iter_rows(min_row=4, values_only=True):
            if row[0] is None:
                continue
            put = {}
            for i, h in enumerate(puts_headers):
                if h is not None:
                    put[h] = row[i]
            open_puts.append(put)
        assigned_headers = [clean_header(cell.value) for cell in assigned_ws[3]]
        for row in assigned_ws.iter_rows(min_row=4, values_only=True):
            if row[0] is None:
                continue
            pos = {}
            for i, h in enumerate(assigned_headers):
                if h is not None:
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
        highest = float(pos['HighestPriceSeen']) if pos.get('HighestPriceSeen') else cost_basis
        if ticker not in all_data or all_data[ticker] is None:
            continue
        current = all_data[ticker]['current']
        static_stop = round(cost_basis * 0.95, 2)
        trailing_active = current >= cost_basis * 1.10 or highest >= cost_basis * 1.10
        trailing_stop = round(highest * 0.95, 2)
        stop_price = trailing_stop if trailing_active else static_stop
        stop_type = 'TRAILING' if trailing_active else 'STATIC'
        pnl_pct = round((current - cost_basis) / cost_basis * 100, 2)
        has_covered_call = pos.get('CoveredCallStrike') is not None
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
            'covered_call_strike': pos.get('CoveredCallStrike'),
            'covered_call_expiry': pos.get('CoveredCallExpiry'),
            'status': status
        })
    return alerts

def get_call_recommendations(assigned, all_data, rules):
    recommendations = []
    for pos in assigned:
        ticker = pos['Ticker']
        cost_basis = float(pos['CostBasis'])
        shares = int(pos['Shares'])
        highest = float(pos.get('HighestPriceSeen') or cost_basis)
        has_covered_call = pos.get('CoveredCallStrike') is not None
        if has_covered_call:
            continue
        if ticker not in all_data or all_data[ticker] is None:
            continue
        current = all_data[ticker]['current']
        trailing_active = highest >= cost_basis * 1.10
        trailing_stop = round(highest * 0.95, 2)
        static_stop = round(cost_basis * 0.95, 2)
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

def build_report(cfg, candidates, earnings_tickers, all_data, watchlist_flags, open_puts, assigned):
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
    committed = 0
    for p in open_puts:
        try:
            committed += float(p['Strike']) * int(p['Contracts']) * 100
        except:
            pass
    assigned_value = 0
    for a in assigned:
        try:
            assigned_value += float(a['CostBasis']) * int(a['Shares'])
        except:
            pass
    total_deployed = committed + assigned_value
    available = total - reserve - total_deployed
    L.append('Total portfolio value:    $' + format(int(total), ','))
    L.append('Dry powder reserve:       $' + format(int(reserve), ','))
    L.append('Cash in open puts:        $' + format(int(committed), ','))
    L.append('Cash in assigned stocks:  $' + format(int(assigned_value), ','))
    L.append('Total deployed:           $' + format(int(total_deployed), ','))
    L.append('Available for new puts:   $' + format(int(available), ','))
    L.append('Note: Update total_value in config.json as it changes.')
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
            prem = p.get('PremiumCollected', 'N/A')
            try:
                prem_str = '$' + str(round(float(prem), 2))
            except:
                prem_str = 'N/A'
            expiry = p.get('Expiry', 'N/A')
            expiry_str = expiry.strftime('%Y-%m-%d') if hasattr(expiry, 'strftime') else str(expiry)
            L.append('Ticker: ' + str(p['Ticker']) + ' | Strike: $' + str(p['Strike']) + ' | Expiry: ' + expiry_str + ' | Contracts: ' + str(p['Contracts']) + ' | Premium: ' + prem_str)
    return '\n'.join(L)

def get_performance_summary():
    try:
        wb = openpyxl.load_workbook('C:\\TradingBot\\positions.xlsx')
        ws = wb['ClosedTrades']
        raw_headers = [cell.value for cell in ws[3]]
        headers = [clean_header(h) for h in raw_headers]
        trades = []
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[0] is None:
                continue
            trade = {}
            for i, h in enumerate(headers):
                if h is not None:
                    trade[h] = row[i]
            trades.append(trade)
        if not trades:
            return None
        total = len(trades)
        expired = len([t for t in trades if t.get('Outcome') == 'EXPIRED'])
        assigned = len([t for t in trades if t.get('Outcome') == 'ASSIGNED'])
        called_away = len([t for t in trades if t.get('Outcome') == 'CALLED_AWAY'])
        stop_losses = len([t for t in trades if t.get('Outcome') == 'STOP_LOSS'])
        btc_count = len([t for t in trades if t.get('Outcome') == 'BTC'])
        results = []
        for t in trades:
            try:
                prem = float(t.get('PremiumCollected') or 0)
                btc_amt = float(t.get('BTCAmount') or 0)
                call = float(t.get('CallPremium') or 0)
                total_income = (prem - btc_amt) + call
                strike = float(t.get('Strike') or 0)
                contracts = int(t.get('Contracts') or 1)
                days_val = t.get('DaysHeld')
                try:
                    days = int(float(str(days_val)))
                except:
                    try:
                        open_date = t.get('OpenDate')
                        close_date = t.get('CloseDate')
                        if open_date and close_date:
                            days = (close_date - open_date).days
                        else:
                            days = 0
                    except:
                        days = 0
                notional = strike * contracts * 100
                ret_pct = total_income / notional if notional > 0 else 0
                results.append({
                    'total_income': total_income,
                    'ret_pct': ret_pct,
                    'prem': prem,
                    'btc_amt': btc_amt,
                    'call': call,
                    'days': days
                })
            except:
                pass
        if not results:
            return None
        incomes = [r['total_income'] for r in results]
        wins = len([i for i in incomes if i > 0])
        win_rate = round(wins / total * 100, 1) if total > 0 else 0
        returns = [r['ret_pct'] for r in results if r['ret_pct'] != 0]
        avg_return = round(sum(returns) / len(returns) * 100, 2) if returns else 0
        best = round(max(returns) * 100, 2) if returns else 0
        worst = round(min(returns) * 100, 2) if returns else 0
        total_premium = round(sum(r['prem'] for r in results), 2)
        total_btc_paid = round(sum(r['btc_amt'] for r in results), 2)
        total_call = round(sum(r['call'] for r in results), 2)
        total_income = round(sum(r['total_income'] for r in results), 2)
        days_list = [r['days'] for r in results if r['days'] > 0]
        avg_days = round(sum(days_list) / len(days_list), 1) if days_list else 0
        return {
            'total': total,
            'expired': expired,
            'assigned': assigned,
            'called_away': called_away,
            'stop_losses': stop_losses,
            'btc': btc_count,
            'win_rate': win_rate,
            'avg_return': avg_return,
            'best': best,
            'worst': worst,
            'total_premium': total_premium,
            'total_btc_paid': total_btc_paid,
            'total_call': total_call,
            'total_income': total_income,
            'avg_days': avg_days
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
        L.append('Total trades:             ' + str(perf['total']))
        L.append('Puts expired worthless:   ' + str(perf['expired']))
        L.append('Puts assigned:            ' + str(perf['assigned']))
        L.append('Covered calls closed:     ' + str(perf['called_away']))
        L.append('Bought to close:          ' + str(perf['btc']))
        L.append('Stop losses:              ' + str(perf['stop_losses']))
        L.append('Win rate:                 ' + str(perf['win_rate']) + '%')
        L.append('Avg return per trade:     ' + str(perf['avg_return']) + '%')
        L.append('Best trade:               ' + str(perf['best']) + '%')
        L.append('Worst trade:              ' + str(perf['worst']) + '%')
        L.append('Total premium collected:  $' + format(perf['total_premium'], ','))
        L.append('Total BTC paid:           $' + format(perf['total_btc_paid'], ','))
        L.append('Total call premium:       $' + format(perf['total_call'], ','))
        L.append('Total net income:         $' + format(perf['total_income'], ','))
        L.append('Avg days held:            ' + str(perf['avg_days']))
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
    print('Fetching stock data for ' + str(len(watchlist)) + ' stocks...')
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
    print('Finding sympathy drop candidates...')
    candidates = find_put_candidates(cfg, all_data, earnings_tickers, open_position_tickers)
    longer_candidates = find_longer_dated_candidates(cfg, all_data, earnings_tickers, open_position_tickers)
    print('Calculating technicals for mean reversion screen...')
    mean_reversion_candidates = find_mean_reversion_candidates(cfg, all_data, earnings_tickers, open_position_tickers)
    print('Building macro calendar...')
    macro_section = build_macro_section(open_puts)
    print('Building report...')
    report = build_report(cfg, candidates, earnings_tickers, all_data, watchlist_flags, open_puts, assigned)
    longer_section = build_longer_dated_section(longer_candidates)
    mean_reversion_section = build_mean_reversion_section(mean_reversion_candidates, earnings_tickers)
    sections_cd = build_sections_cd(open_puts, assigned, all_data, rules)
    perf_section = build_performance_section()
    full_report = (report + '\n' +
                   longer_section + '\n' +
                   mean_reversion_section + '\n' +
                   sections_cd + '\n' +
                   perf_section + '\n' +
                   macro_section)
    print('')
    print(full_report)
    print('')
    print('Sending email...')
    today = datetime.now().strftime('%Y-%m-%d')
    send_email(cfg, 'Trading Bot Morning Report - ' + today, full_report)
    print('Report emailed successfully.')

if __name__ == '__main__':
    main()

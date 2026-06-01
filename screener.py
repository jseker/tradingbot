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
        expired = len([t for t in trades if t.get('Outcome') == 'EXPIRED'])
        assigned = len([t for t in trades if t.get('Outcome') == 'ASSIGNED'])
        called_away = len([t for t in trades if t.get('Outcome') == 'CALLED_AWAY'])
        stop_losses = len([t for t in trades if t.get('Outcome') == 'STOP_LOSS'])
        btc = len([t for t in trades if t.get('Outcome') == 'BTC'])
        results = []
        for t in trades:
            try:
                prem = float(t.get('PremiumCollected') or 0)
                btc_amt = float(t.get('BTCAmount') or 0)
                call = float(t.get('CallPremium') or 0)
                strike = float(t.get('Strike') or 0)
                contracts = int(t.get('Contracts') or 1)
                days = int(t.get('DaysHeld') or 0)
                notional = strike * contracts * 100
                total_income = (prem - btc_amt) + call
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
            'btc': btc,
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

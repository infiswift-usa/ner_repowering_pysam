**data types**
self.reference_prices: dict of regions and last column values {region : latest year's avg}
self.non_fossil_value: float value of last input of jepx 加重平均値
self.balancing_costs: list of col[1] (since it's currently used in calculation).eg., [2,1.8,1.6,1.4,1.3,1.3,1.3,1.3,1.3,1.3,0.3...]
self.ppa_prices:dict of regions and respective prices {region:ppa_price}

**VARIABLES**

*FIT残月数* --> FORMULA  =20*12-DATEDIF(C15,C16,"M")
in our code:
    fit_remaining_months = 20 * 12 - self.month_diff(params['op_start_date'], params['mod_date'])

*a.基準(按分)価格*
base_price_a = (params['fit_price'] * params['ex_dc'] + params['latest_price'] * (params['rep_dc'] - params['ex_dc'])) / params['rep_dc']

*b.参照価格:*
ref_price_b = self.reference_prices.get(params['region'], 0.0)

*c.非化石価値相当額:*
non_fossil_val_c = params.get('non_fossil_value', self.non_fossil_value)

*⑤PPA単価:*
ppa_price = self.ppa_prices.get(params['region'], 14.0)

**TABLE**
*column[0] is year. starting from 0*

*column[1] of price prediction table:*  *d.バランシングコスト:* 
    bal_cost_d = self.balancing_costs[year - 1] if (year - 1) < len(self.balancing_costs) else 0.30 
    *->* year is in range of 1,21. so [2,1.8,1.6,1.4,1.3,1.3,1.3,1.3,1.3,1.3,0.3...] is accessed from index 0.

*column[2]: ①FIPプレミアム*
    =a-b-c+d
    
*column[3]: 売電単価*
    =IF(G19*12<=$I$2,$I$12+I19,IF((G19-1)*12<$I$2,($I$12+I19)/12*($I$2-(G19-1)*12)+$I$12/12*(G19*12-$I$2),$I$12))
    *we do:*
    months_at_end_of_year = year * 12
    months_at_start_of_year = (year - 1) * 12
    *fit_remaining_months is FIT残月数*
    if months_at_end_of_year <= fit_remaining_months: 
        sell_price = ppa_price + fip_premium
    elif months_at_start_of_year < fit_remaining_months:
        fip_months = fit_remaining_months - months_at_start_of_year
        non_fip_months = months_at_end_of_year-fit_remaining_months
        sell_price = ((ppa_price + fip_premium) * fip_months + ppa_price * non_fip_months) / 12.0
    else:
        sell_price = ppa_price

*column[4]:年間発電量*
     gen_kwh = params['rep_yield'] for row 1. for consecutive rows, it is =K19*(1-$E$9) *-->* gen_kwh = gen_kwh * (1.0 - params['rep_deg'])
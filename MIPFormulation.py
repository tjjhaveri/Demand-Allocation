import pandas as pd
from pulp import *

def Function(raw_df,cat_preferred_Max=0.7,cat_not_preferred_Max=0.2,utilization_score_Min=0.95,lead_time_importance_level='Medium'):
    raw_df = raw_df.astype(dict(zip(raw_df.columns,[str,str,str,str,'float64','float64','float64','float64'])))
    raw_df['Capacity_Quantity'] = raw_df['Capacity']*raw_df['Tesla_Demand']
    #remove banned supplier
    df = raw_df[raw_df['Supplier_Preference']!='ban']
    #extract single sourced
    max_vendor_number_df = df.copy()[['Vendor_Name','TPN']].groupby('TPN').count().reset_index()
    single_sourced = list(max_vendor_number_df[max_vendor_number_df['Vendor_Name']==1]['TPN'])
    single_sourced_df = df[df['TPN'].isin(single_sourced)]
    multi_sourced = list(max_vendor_number_df[max_vendor_number_df['Vendor_Name']>1]['TPN'])
    df = df[df['TPN'].isin(multi_sourced)]
    print(single_sourced_df)


    # define this as a minimization problem
    prob = LpProblem("Component_Allocation",LpMinimize)
    #solver = pulp.PULP_CBC_CMD(keepFiles=True)
    #lead time importance
    lead_time_importance_dict = {'High': 100.0, 'Medium': 5.0, 'Low': 1.0}
    lead_time_importance = lead_time_importance_dict[lead_time_importance_level]
    #vendor list and part list
    vendors = list(df.Vendor_Name.unique())
    parts = list(df.TPN.unique())
    categories = list(df.Component_Category.unique())
    #production cost data formating
    cost_df = pd.pivot_table(df[['Vendor_Name','TPN','Pricing_per_Item']],values='Pricing_per_Item',index='TPN',columns='Vendor_Name').fillna(0.0)
    cost = {parts[i]:{vendors[j]:cost_df.loc[parts[i],vendors[j]] for j in range(len(vendors))} for i in range(len(parts))}
    #production capacity data formating
    production_capacity_df =  pd.pivot_table(df[['Vendor_Name','TPN','Capacity_Quantity']],values='Capacity_Quantity',index='TPN',columns='Vendor_Name').fillna(0.0)
    production_capacity = {parts[i]:{vendors[j]:production_capacity_df.loc[parts[i],vendors[j]] for j in range(len(vendors))} for i in range(len(parts))}
    #lead time data formating
    lead_time_df = df.copy()[['Vendor_Name','Lead_Time_wks']].drop_duplicates()
    lead_time = dict(zip(lead_time_df['Vendor_Name'],lead_time_df['Lead_Time_wks']))
    #demand data formating
    demand_df = df.copy()[['TPN','Tesla_Demand']].drop_duplicates()
    demand = dict(zip(demand_df['TPN'],demand_df['Tesla_Demand']))
    #supplier score data formating
    supplier_score_df = df[['Vendor_Name','Component_Category','Supplier_Preference']].groupby(['Vendor_Name','Component_Category','Supplier_Preference']).sum().reset_index()
    supplier_score = {cat: {} for cat in categories}
    for cat in categories:
        temp = supplier_score_df[supplier_score_df['Component_Category']==cat]
        supplier_score[cat]=dict(zip(temp['Vendor_Name'],temp['Supplier_Preference']))
    #max vendor numbers data formating
    max_vendor_number_df[max_vendor_number_df['Vendor_Name'] > 6] = 6
    max_vendor_number = dict(zip(max_vendor_number_df['TPN'],max_vendor_number_df['Vendor_Name']))

    # define decision variables
    allocation = [(part,vendor) for part in parts for vendor in vendors]
    usage_control = [(part,vendor) for part in parts for vendor in vendors]
    allocation_vars = LpVariable.dicts("Allocation",allocation,lowBound=0,cat='Continuous')
    usage_control_vars = LpVariable.dicts("Usage_Control",usage_control,cat='Binary')

    # objective function
    prob += lpSum([allocation_vars[(part,vendor)]*cost[part][vendor]*demand[part] for part in parts for vendor in vendors])\
            +lpSum([usage_control_vars[(part,vendor)]*(lead_time[vendor]**0.5)*cost[part][vendor]*1.65*(demand[part]/52)\
                    *lead_time_importance for part in parts for vendor in vendors])

    # s.t.
    # 1
    for part in parts:
        for vendor in vendors:
            prob += allocation_vars[(part, vendor)] * demand[part] <= production_capacity[part][vendor]

    # 2
    for part in parts:
        for vendor in vendors:
            cap = min(1 / (1 + max_vendor_number[part]), production_capacity[part][vendor] / demand[part])
            prob += allocation_vars[(part, vendor)] >= cap
            #prob += allocation_vars[(part, vendor)] >= 0.2

    # 3.
    for part in parts:
        prob += lpSum([allocation_vars[(part, vendor)] for vendor in vendors]) >= 1.0

    # 4.
    for part in parts:
        # no more than Ni source
        prob += lpSum([usage_control_vars[(part, vendor)] for vendor in vendors]) <= max_vendor_number[part]
        # use at least 2 vendors
        prob += lpSum([usage_control_vars[(part, vendor)] for vendor in vendors]) >= 2

    # 5.
    # Big M operation
    M = 10 ** 10
    # link variables
    for part in parts:
        for vendor in vendors:
            prob += usage_control_vars[(part, vendor)] * M >= allocation_vars[(part, vendor)]
            prob += usage_control_vars[(part, vendor)] <= allocation_vars[(part, vendor)] * M

    # 6.
    # find exemption part and exclude from part list
    not_high_list = set(df[df.Supplier_Preference != 'high']['TPN'].unique())
    exclude_part = list(not_high_list.difference(set(parts)))

    for cat in categories:
        print(cat)
        all_temp_part_list = list(df[df.Component_Category == cat]['TPN'])
        temp_vendor_list = list(df[df.Component_Category == cat]['Vendor_Name'].unique())
        temp_part_list = [part for part in all_temp_part_list if part not in exclude_part]
        for part in temp_part_list:
            for vendor in temp_vendor_list:
                if supplier_score[cat][vendor] == 'low':
                    prob += lpSum([usage_control_vars[(part, v)] for v in vendors if supplier_score[cat][v] == 'high']) >= \
                            usage_control_vars[(part, vendor)]


    # 7.
    for cat in categories:
        temp_part_list = list(df[df.Component_Category == cat]['TPN'])
        temp_vendor_list = list(df[df.Component_Category == cat]['Vendor_Name'].unique())
        for vendor in temp_vendor_list:
            if supplier_score[cat][vendor] == 'high':
                prob += lpSum([allocation_vars[(part, vendor)] * demand[part] for part in temp_part_list]) <= cat_preferred_Max * lpSum(
                    [demand[part] for part in temp_part_list])
            else:
                prob += lpSum([allocation_vars[(part, vendor)] * demand[part] for part in temp_part_list]) <= cat_not_preferred_Max * lpSum(
                    [demand[part] for part in temp_part_list])

    # 8.
    for part in parts:
        for vendor in vendors:
            prob += allocation_vars[(part, vendor)] <= 1
            prob += allocation_vars[(part, vendor)] >= 0

    # 9.
    prob += lpSum([usage_control_vars[(part, vendor)] for vendor in vendors for part in parts]) / lpSum(
        [max_vendor_number[part] for part in parts]) >= utilization_score_Min



    prob.solve() #sepecify solver if needed
    if prob.solve()==1:
        print('Optimal solution found:')
        obj = value(prob.objective)
        print("         The total cost of this allocation is: ${}".format(round(obj,3)))
        sourcing_usage = sum([usage_control_vars[(part, vendor)].varValue for part in parts for vendor in vendors])/sum(max_vendor_number.values())
        print('         {:.2f}% of sources are utilized.'.format(sourcing_usage*100))
        print('')
    else:
        print('Failed to find optimal solution.')

    print(df)
    # construct output df
    output = []
    for cat in categories:
        for part, vendor in allocation_vars:
            print(part,vendor)
            try:
                if cost[part][vendor] > 0:
                    var_output = {
                        'Telsa PN': part,
                        'Vendors': vendor,
                        '2020 Price': round(cost[part][vendor], 5),
                        'Lead Time': lead_time[vendor],
                        'Supplier Preference': supplier_score[cat][vendor],
                        'Capacity': round(production_capacity[part][vendor] / demand[part], 3),
                        'Allocation Fraction': round(allocation_vars[(part, vendor)].varValue, 4),
                        'Allocation Quantity': int(round(allocation_vars[(part, vendor)].varValue * demand[part], 0))
                    }
                output.append(var_output)
            except:
                pass
    output_df = pd.DataFrame.from_records(output).sort_values(['Telsa PN', 'Vendors'])
    output_df.set_index(['Telsa PN', 'Vendors'], inplace=True)
    output_df = output_df.reset_index()
    print(single_sourced_df.columns)
    single_sourced_df = single_sourced_df[['TPN','Vendor_Name','Pricing_per_Item','Lead_Time_wks','Supplier_Preference','Capacity','Capacity','Tesla_Demand']]
    single_sourced_df.columns = ['Telsa PN', 'Vendors', '2023 Price', 'Lead Time', 'Supplier Preference', 'Capacity',\
       'Allocation Fraction', 'Allocation Quantity']
    output_df = pd.concat([output_df,single_sourced_df])
    return output_df

d = pd.read_excel('Data_with_Forecast.xlsx', engine='openpyxl', sheet_name= 'PCB - Chris')
output = Function(d)

output.to_excel("Allocation_Output_Revised_PCB.xlsx")

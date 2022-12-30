import datetime
import os
import pandas as pd
import json
import numpy as np
#import plotly
#import plotly.express as px
import plotly.graph_objects as go


test_name =['damp-heat','light-cycle','mpp'] 
damp_heat= 'damp-heat'
light_cycle= 'light-cycle'
mpp= 'mpp'
batch_name='old'


def read_modules(folder):
    '''
    read all xlsm name
    '''
    filename_list=[]
    for root, ds, fs in os.walk(folder):
        for f in fs:
            filename_list.append(f)
    return filename_list


def read_cell_pce(file_path,filename):
    '''
    read cell performance in one xlsm
    '''
    file_path = file_path+filename
    df = pd.read_excel(file_path, sheet_name = "Sheet2")
    df.dropna(axis=1,inplace=True)
    df.drop(index=0,inplace=True)
    df.drop(['Hystersis'],axis=1,inplace=True)
    
    df['name'] = pd.to_datetime(df['name'])
    start_time= df.at[1, 'name']
    start_state = df.iloc[0]
    data_list = []
    ratio_list= []
    max_pce=[]
    max_pce_ratio=[]
    all_time=set()
    for index, row in df.iterrows():
        hour_data= []
        hour_data_ratio=[]
        hours = (row['name']-start_time).total_seconds()/3600
        hour_data.append(hours)
        all_time.add(hours)
        hour_data_ratio.append(hours)
        hour_data_ratio.append(row['name'])
        for i, v in row.iteritems():
            hour_data.append(v)
            if i == 'PCE_r':
                max_pce.append((hours,v))
                max_pce_ratio.append((hours,v/start_state[i]))
            if i!='name':
                hour_data_ratio.append(v/start_state[i])     
        data_list.append(hour_data)
        ratio_list.append(hour_data_ratio)
    
    filename = filename[:-5]
    if 'UV' in filename:
        filename = filename.replace('UV','')+" with UV filter"
    data ={'test_type':file_path.split('/')[2],'cell_name':filename,'cell_data':data_list,'data_ratio':ratio_list,'max_pce':max_pce,'max_pce_ratio':max_pce_ratio,'all_time':sorted(all_time)}
    return(data)


def read_all_test():
    '''
    generate all  cell performance from all xlsm files
    '''
    all_test_data = []
    for i in test_name:
        file_path='./'+batch_name+'/'+i+'/'
        name_list = read_modules(file_path)
        for j in name_list:
            all_test_data.append(read_cell_pce(file_path,j))
    return(all_test_data)


def draw_cell_PCE_overview(data_list):
    '''
    draw all cell PCE data
    '''
    categories = []
    for i in data_list:
        categories.extend(i['all_time'])
    categories = sorted(set(categories))
    fig = go.Figure()
    
    for i in data_list:
        cell_array = [np.nan]*len(categories)
        for j in i['max_pce']:
            cell_array[categories.index(j[0])] = j[1]
        fig.add_trace(
            go.Scatter(
                x = categories,
                y = cell_array,
                name = i['test_type']+': '+i['cell_name'],
                line_shape = 'linear',
                mode = 'lines+markers',
                connectgaps=True
            )
        )
    fig.update_layout(title='PCE', 
                      xaxis = dict(
                          tickmode = 'array',
                          tickvals = categories
                      ))    
    fig.show()
    

def draw_cell_PCE_overview_ratio(data_list):
    '''
    draw all cell PCE ratio
    '''
    categories = []
    for i in data_list:
        categories.extend(i['all_time'])
    categories = sorted(set(categories))
    fig = go.Figure()
    
    for i in data_list:
        cell_array = [np.nan]*len(categories)
        for j in i['max_pce_ratio']:
            cell_array[categories.index(j[0])] = j[1]
        fig.add_trace(
            go.Scatter(
                x = categories,
                y = cell_array,
                name = i['test_type']+': '+i['cell_name'],
                line_shape = 'linear',
                mode = 'lines+markers',
                connectgaps=True
            )
        )
    fig.update_layout(title='PCE Ratio',
                      xaxis = dict(
                          tickmode = 'array',
                          tickvals = categories
                      ))    
    fig.show()
    
    
def draw_cell_performance(data_list):
    '''
    draw 'Voc_r','Jsc_r','FF_r','PCE_r','Voc_f','Jsc_f','FF_f','PCE_f' from each cell
    '''
    categories = ['Voc_r','Jsc_r','FF_r','PCE_r','Voc_f','Jsc_f','FF_f','PCE_f']
    for i in data_list:
        fig = go.Figure()
        for j in i['data_ratio']:
            data = j[2:]
            fig.add_trace(go.Scatterpolar(
            r=data,
            theta=categories,
            name=str(j[0])
            ))
        fig.update_layout(
        polar=dict(
        radialaxis=dict(
        visible=True,
        title=i['test_type'] + ': '+i['cell_name']
        )),
        showlegend=True
        )
        fig.show()


def read_performance(data_list):
    cell_performance=dict()
    cell_performance_ratio = dict()
    cell_PCE_overview=dict()
    cell_PCE_overview_ratio = dict()
    for i in data_list:
        cell_performance = []
        overview_performance = dict()
        overview_performance_ratio = dict()
        for j in i['cell_data']:
            cell_performance_per_hours=dict()
            cell_performance_per_hours['Voc_r'] = j[2]
            cell_performance_per_hours['Jsc_r'] = j[3]
            cell_performance_per_hours['FF_r'] = j[4]
            cell_performance_per_hours['PCE_r'] = j[5]
            cell_performance_per_hours['Voc_f'] = j[6]
            cell_performance_per_hours['Jsc_f'] = j[7]
            cell_performance_per_hours['FF_f'] = j[8]
            cell_performance_per_hours['PCE_f'] = j[9]
            cell_performance[j[0]] = cell_performance_per_hours
        for j in i['data_ratio']:
            cell_performance_per_hours=dict()
            cell_performance_per_hours['Voc_r'] = j[2]
            cell_performance_per_hours['Jsc_r'] = j[3]
            cell_performance_per_hours['FF_r'] = j[4]
            cell_performance_per_hours['PCE_r'] = j[5]
            cell_performance_per_hours['Voc_f'] = j[6]
            cell_performance_per_hours['Jsc_f'] = j[7]
            cell_performance_per_hours['FF_f'] = j[8]
            cell_performance_per_hours['PCE_f'] = j[9]
            cell_performance_ratio[j[0]] = cell_performance_per_hours
        for j in i['max_pce']:
            overview_performance[j[0]] = j[1]
        for j in i['max_pce_ratio']:
            overview_performance_ratio[j[0]] = j[1]
        cell_performance[i['cell_name']] = overview_performance
        cell_performance_ratio[i['cell_name']] = overview_performance_ratio
        cell_PCE_overview[i['cell_name']] =  overview_performance
        cell_PCE_overview_ratio[i['cell_name']] =  overview_performance_ratio


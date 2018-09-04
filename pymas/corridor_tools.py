import re
import os
import numpy as np 
import pandas as pd 
import zipfile
import pickle
import openpyxl

# Plotting
import matplotlib # '2.1.2'
from matplotlib.ticker import NullLocator
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
#------------------------------------------------------------------------------
def get_inputs(step=None, interface='0_Interface.xlsx', sheet='Inputs'):
    
    xl_f = interface
    df = pd.read_excel(xl_f, sheet_name=sheet, header=None)
    
    inputs = {}
    
    # General inputs:
    inputs['ccr'] = df.iloc[2,1]
    inputs['base_year'] = df.iloc[3,1]
    inputs['curr_year'] = df.iloc[4,1]
    inputs['analyst'] = df.iloc[5,1]
    
    # 2 Truck Percentages
    if str(step) == '2':
        inputs['yrs'] = list(df.iloc[2:7,3].dropna())
        
    # 3a Corridor Analysis
    elif str(step) == '3a':
        sfx = str(df.iloc[2,5])
        inputs['suffix'] = '' if sfx == 'nan' else sfx
        inputs['xl_vars'] = list(df.iloc[2:7, 7].dropna())
        inputs['hm_vars'] = list(df.iloc[2:7, 9].dropna())
        
        norm = df.iloc[2,11]
        inputs['hm_norm'] = None if str(norm)=='nan' else norm
        inputs['update_pp'] = df.iloc[6, 11]
        inputs['line_vars'] = list(df.iloc[2:7, 13].dropna())  
    
    # 3b Corridor Analysis (loops)
    elif str(step) == '3b':
        sfx_i = str(df.iloc[11,5])
        inputs['suffix_in'] = '' if sfx_i == 'nan' else sfx_i
        
        sfx_o = str(df.iloc[12,5])
        inputs['suffix_out'] = '' if sfx_o == 'nan' else sfx_o
        
        inputs['update_pp'] = df.iloc[11, 7]
        inputs['pctgd'] = df.iloc[11, 9]
        inputs['by_day_fmt'] = df.iloc[11, 11]
        inputs['by_mp_fmt'] = df.iloc[11, 13]
        
    # 4a Commute Setup
    elif str(step) == '4a':
        inputs['source'] = df.iloc[2, 16]
        inputs['cleanup_fname'] = df.iloc[2, 18]
    
    # 4b Commute Analysis
    elif str(step) == '4b':
        sfx_i = str(df.iloc[11, 16])
        inputs['suffix_in'] = '' if sfx_i == 'nan' else sfx_i
        
        sfx_o = str(df.iloc[12, 16])
        inputs['suffix_out'] = '' if sfx_o == 'nan' else sfx_o
        
        p_types = df.iloc[11:16, 18].dropna()
        if len(p_types) == 0:
            inputs['plot_types'] = ['hov_cong', 'gp_cong', 'gp_sc', 
                                    'gp_spd', 'gp_tt'] 
        else:
            inputs['plot_type'] = p_types
            
    return inputs

#------------------------------------------------------------------------------
def get_batchlist(analysis='corridor', interface='0_Interface.xlsx', 
                  sheet='Batch'):
    
    xl_f = interface
    df = pd.read_excel(xl_f, sheet_name=sheet)
    
    if analysis == 'corridor':
        df = df.iloc[1:,:]
        res = {}
        for reg in df.columns:
            nans = df[df[reg].map(lambda x: str(x) == 'nan')]
            res[reg] = list(nans.index.values)
    
    elif analysis == 'commute':
        df = df.iloc[0, :]
        res = list(df[df.map(lambda x: str(x)=='nan')].index)
    else:
        raise Error("analysis must be 'commute' or 'corridor'")
    return res

#------------------------------------------------------------------------------
def read_object(name, obj_class, paths, suffix=''):
    
    path_key = {
        'Region' : 'com_out_path',
        'Corridor' : 'cor_out_path',
        'LoopGroup' : 'loop_path'
    }
    
    path = paths[path_key[obj_class]]
    
    fname = '%s%s.dat'%(name, suffix)
    
    with open(os.path.join(path, fname), 'rb') as f:
        obj = pickle.load(f)
        
    return obj
    
#------------------------------------------------------------------------------
def define_paths(ccr, region, base_year, curr_year):

    default_paths = {
        #Setup/config
        'cor_cfg_path' : './%s/0_Inputs/%s/'%(ccr, region),
        'sw_cfg_path' : './%s/0_Inputs/'%(ccr),
        'tmp_path' : './%s/0_Inputs/_Templates/'%(ccr),
        'tk_pct_path' : './%s/0_Inputs/_TruckPercent/'%(ccr),
        
        #Inputs/Data
        'hist_cor_path' : './%s/1_Data/%s/1_Corridor Data'%(ccr, 
                                                            region),
        
        'base_cont_path' : './%s/1_Data/%s/1_Corridor Data/%s'%(ccr, 
                                                                region, 
                                                                base_year),
        'current_cont_path' : './%s/1_Data/%s/1_Corridor Data/%s'%(ccr, 
                                                                   region, 
                                                                   curr_year),
        'base_com_path' : './%s/1_Data/%s/2_Commute Data/%s'%(ccr, 
                                                              region,
                                                              base_year),
        
        'current_com_path' : './%s/1_Data/%s/2_Commute Data/%s'%(ccr, 
                                                              region,
                                                              curr_year),
        
        'com_xl_path' : './%s/0_Inputs/%s/'%(ccr, region),
        
        'throughput_path' : './%s/1_Data/%s/3_Throughput Data/'%(ccr, 
                                                                 region),
        
        'base_loop_path' : './%s/1_Data/%s/4_Loop Data/%s'%(ccr, 
                                                            region, 
                                                            base_year),
        
        'current_loop_path' : './%s/1_Data/%s/4_Loop Data/%s'%(ccr, 
                                                               region, 
                                                               curr_year),
        'loop_path' : './%s/1_Data/%s/4_Loop Data/'%(ccr, region),
        
        #Outputs
        'cor_out_path' : './%s/2_Corridor Output/%s/'%(ccr, region),
        
        'com_out_path' : './%s/3_Commute Output/%s/'%(ccr, region),
        
        #Plotting
        'plot_path' : './%s/4_Plots/%s/'%(ccr, region)

    }
    
    default_paths = pd.Series(default_paths)
        
    # Get Filepaths from Interface file.

    paths = pd.read_excel('0_Interface.xlsx',
                          sheet_name = 'Filepaths')

    # filter to region
    paths = paths[region]
    paths.index = paths.index.map(lambda x: x + '_path')
    replace = paths.map(lambda x: str(x) == 'nan')
    paths[replace] = default_paths[replace]

    return paths

#------------------------------------------------------------------------------
def folder_setup(ccr, yrs):

    # Directory Structure
    level1 = ['0_Inputs', '1_Data', '2_Corridor Output', 
             '3_Commute Output', '4_Plots']

    # Second level for all level1 directories are the regions:
    regions = ['NWR', 'OR', 'SWR', 'ER', 'SCR']

    # Subdirectories for each region in '1_Data' directory
    data_types = ['1_Corridor Data', '2_Commute Data',
                  '3_Throughput Data', '4_Loop Data']


    yrs = map(str, yrs)

    if not os.path.exists(ccr):
        os.mkdir(ccr)

    # First level of folders under 'CCR [year]'
    for level_1 in level1:

        # Create folders if they do not exist
        if not os.path.exists(os.path.join(ccr, level_1)):
            os.mkdir(os.path.join(ccr, level_1))
        
        
        # Create next level of folders; one for each region
        for region in regions:
            if not os.path.exists(os.path.join(ccr, level_1, region)):
                os.mkdir(os.path.join(ccr, level_1, region))

            if level_1 == '1_Data':
                
                # Create subfolders for each input data type
                for data_type in data_types:
                    if not os.path.exists(os.path.join(ccr, level_1, 
                                                       region, data_type)):
                        
                        os.mkdir(os.path.join(ccr, level_1, region, data_type))
                
                    # For corridor and commute data, create year subfolders
                    if data_type in ['1_Corridor Data', '2_Commute Data',
                                     '4_Loop Data']:               
                        for yr in yrs:    
                            if not os.path.exists(os.path.join(ccr, level_1, 
                                                               region,
                                                               data_type, 
                                                               yr)):
                                
                                os.mkdir(os.path.join(ccr, level_1, region, 
                                                      data_type, yr))

#------------------------------------------------------------------------------
def pct_change(base, current):
    base = float(base)
    current = float(current)
    return np.nan if base == 0 else (current - base) / base

# vectorized version of pct_change
def vect_pct_change(base_df, current_df):
    vpc = np.vectorize(pct_change)
    
    pct = vpc(base_df, current_df)
    
    if len(pct.shape) == 2:   #DataFrame
        return pd.DataFrame(pct, 
                            index=base_df.index, 
                            columns=base_df.columns)
    else:  #Series
        return pd.Series(pct, index=base_df.index)

#------------------------------------------------------------------------------
def calc_delay(spd, vol, sl, thr):
    '''Returns the delay (veh-hours per mile) from threshold 
    speed (thr * 60).
    
    Parameters:
    spd - speed
    vol - volume (in number of vehicles in time period, or as vmt)
    sl - speed limit
    thr - delay threshold (percent)'''
    
    if spd > 0 and spd < thr * sl:
        diff = (1./spd) - (1./(thr*sl))
    elif spd > thr * sl:
        diff = 0.
    else:
        diff = np.nan
    
    delay = diff * vol
    
    return delay

# vectorized version of calc_delay
def vect_delay(spd_df, vol_df, sl, thr):
    vd = np.vectorize(calc_delay)
    with np.errstate(divide='ignore',invalid='ignore'):
        delay_array = vd(spd_df, vol_df, sl, thr)

    if len(delay_array.shape) == 2:   #DataFrame
        return pd.DataFrame(delay_array, 
                            index=spd_df.index, 
                            columns=spd_df.columns)
    else:  #Series
        return pd.Series(delay_array, index=spd_df.index)    

    
#------------------------------------------------------------------------------
def unzip_loops(cor, paths):
    for b_c in ['base', 'current']:
        path = paths['%s_loop_path'%b_c]
        path_zf = os.path.join(path, 'ZipFiles')
        path_pf = os.path.join(path, 'Plots')    
        
        # Create ZipFiles and Plots folders
        if not os.path.exists(os.path.join(path_zf)):
                os.mkdir(os.path.join(path_zf))
        if not os.path.exists(os.path.join(path_pf)):
                os.mkdir(os.path.join(path_pf))    
                
        # loop through zip files
        for z in [x for x in os.listdir(path) if '.zip' in x]:

            # extract zip files
            print('Extracting ' + z)

            zf = os.path.join(path, z)
            zf_f = zipfile.ZipFile(zf)
            zf_f.extractall(path)
            zf_f.close()

            # move zip files into 'Zipfiles'
            zf_move = os.path.join(path_zf, z)
            os.rename(zf, zf_move)

        # move png files into 'Plots'
        for f in [f for f in os.listdir(path) if '.png' in f]:
            pf = os.path.join(path, f)
            pf_move = os.path.join(path_pf, f)

            os.rename(pf, pf_move)        
        
#------------------------------------------------------------------------------
def get_loops(cor, paths):
    """Returns a list of loopfiles for a given corridor (cor)"""
    res = []
    
    for b_c in ['base', 'current']:
        p = paths['%s_loop_path'%b_c]
        loops = [x.split()[0] for x in os.listdir(p) if \
                 re.match('0{,2}%ses\d{5}'%cor,x) and x.endswith('.xlsx')]
        res = res + loops
    res = sorted(list(set(res)))
    
    return res

#------------------------------------------------------------------------------
def truck_percentage(ccr, yrs):
    
    if not isinstance(yrs, list):
        yrs = [yrs]
    yrs = map(int, yrs)

    reg_cors = get_batchlist()

    # iterate through corridors dictionary
    for region, cor_list in reg_cors.iteritems():
        cfg_path, pct_path = define_paths(ccr, region, '', '')\
                                 [['cor_cfg_path', 'tk_pct_path']].tolist()

        # loop through corridors for each region
        for cor in cor_list:

            # read config file into dictionary
            xl_obj = pd.ExcelFile(os.path.join(cfg_path,
                                               '%s_%s_config.xlsx'%(cor, 
                                                                    region)))

            cfg_dfs = dict()

            for sheet in xl_obj.sheet_names:
                cfg_dfs[sheet] = pd.read_excel(xl_obj, sheet_name = sheet)

            if isinstance(cfg_dfs[str(yrs[0])].index[0], basestring):
                continue

            # Loop through years
            for yr in yrs:
                # Create dataframe
                config = cfg_dfs[str(yr)]

                # Define list of 0, 1, and 2 years back to compute average
                back_yrs = [yr - yr2 for yr2 in [0,1,2]]

                # Loop through back years and retrieve truck percentages
                for yr2 in back_yrs:

                    # Read TruckPercent csv file
                    df = pd.read_csv(os.path.join(pct_path,
                                                  "TruckPercent%s.csv"%(yr2)))
                    df.loc[:,"Route_ID"] = df.loc[:,"Route_ID"].map(str)

                    # Select main routes (i.e. no suffix in Route_ID)
                    df = df.loc[df.loc[:,"Route_ID"].map(len) < 4,:]

                    # Extract beginning milepost from "Location" variable
                    df.loc[:,'BegMP'] = df.Location.str.extract('(\d{0,3}\.\d{2})',
                                                               expand=False)
                    df.loc[:,'BegMP'] = pd.to_numeric(df.loc[:,'BegMP'])

                    # Group dataframe by sections with same truck percentages
                    df = df.groupby(by=['Route_ID', 
                                        'SingleUnitTruckPct', 
                                        'DoubleUnitTruckPct', 
                                        'TripleUnitTruckPct']). \
                                            agg({'BegMP' : 'min'}) 
                    df = df.reset_index()
                    df['Route_ID'] = pd.to_numeric(df['Route_ID'])

                    # Adjust percentages
                    df['ST'] = df['SingleUnitTruckPct']/100.
                    df['CT'] = (df['DoubleUnitTruckPct'] + 
                                df['TripleUnitTruckPct'])/100.

                    # Select current corridor and sort by milepost
                    df = df[df.loc[:,'Route_ID'] == cor]
                    df = df.sort_values('BegMP').reset_index(drop = True)

                    # Loop through config mileposts 
                    for i, row in config.iterrows():
                        row_assigned = False
                        # assign percentage to each milepose
                        for j, row_lkp in df.iterrows():
                            if j + 1 == len(df):
                                if not row_assigned:
                                    config.loc[i, 'ST%s'%(yr2)] = df.loc[j, 'ST']
                                    config.loc[i, 'CT%s'%(yr2)] = df.loc[j, 'CT'] 
                            else:
                                if (i >= df.loc[j,'BegMP'] and 
                                        i <= df.loc[j+1, 'BegMP']):
                                    config.loc[i,'ST%s'%(yr2)] = df.loc[j, 'ST']
                                    config.loc[i, 'CT%s'%(yr2)] = df.loc[j, 'CT']
                                    row_assigned = True

                # Compute three year rolling averages
                config['ST_pct'] = (config['ST%s'%(back_yrs[0])] + 
                                     config['ST%s'%(back_yrs[1])] + 
                                     config['ST%s'%(back_yrs[2])])/3

                config['CT_pct'] = (config['CT%s'%(back_yrs[0])] + 
                                     config['CT%s'%(back_yrs[1])] + 
                                     config['CT%s'%(back_yrs[2])])/3

                # Drop previous year columns
                cfg_dfs[str(yr)] = config.drop(['ST%s'%(b_y) for b_y in back_yrs]+ 
                                               ['CT%s'%(b_y) for b_y in back_yrs], 
                                               axis = 1)
            try:
                # Write to Excel and save
                xl_f = os.path.join(cfg_path,'%s_%s_config.xlsx'%(cor, region))
                xl_wrt = pd.ExcelWriter(xl_f)

                for sheet in xl_obj.sheet_names:
                    idx=False if sheet == 'Plotting' else True
                    cfg_dfs[sheet].to_excel(xl_wrt, sheet_name=sheet, index=idx)

                xl_wrt.save()

            except IOError:
                print("Could not write %s"%(xl_f))

#------------------------------------------------------------------------------
def commute_setup(yrs, region, ccr, paths, source='Both', cleanup_fname=True):
    
    if source in ['_Commutes.xlsx', 'Both']:
        xl_f = os.path.join(paths['com_xl_path'], '%s_Commutes.xlsx'%region)
        df = pd.read_excel(xl_f, sheet_name='commutes')        
        com_dirs = df.iloc[1,3:].dropna().tolist()
        def_dirs = [paths['base_com_path'], paths['current_com_path']]
        directories = com_dirs + def_dirs if source=='Both' else com_dirs
    
    elif source == 'Default':
        directories = [paths['base_com_path'], paths['current_com_path']]

    else: 
        raise Error()

    trac_df = pd.DataFrame(columns = ["Filename", "Length", "Type", "N GD"])
    
    # create lists of files in the directories and output to .csv
    for drct in directories:
        print('\nSearching %s'%drct)
        if cleanup_fname:
            for yr in yrs:
                # create "remove" pattern from year and rename files 
                remove = '.{0}-01-01.{01}-01-01'.format(str(yr), str(yr + 1))
                from_list = [os.path.join(drct, x) for x in os.listdir(drct)\
                             if remove in x]

                for f in from_list:
                    os.rename(f, f.replace(remove, ''))

        # list Excel files in directory
        trac_list = [x for x in os.listdir(drct) if '.xlsx' in x]

        for com in trac_list:
            name = com

            xl_file = pd.ExcelFile(os.path.join(drct, com))


            # if no "Trip Metadata" sheet, then it is not a commute file
            try:
                df_md = xl_file.parse(sheet_name = 'Trip Metadata', 
                                      header = None)

            except:
                continue

            print('Summarizing %s'%(com))        
            length = df_md.iloc[1,1]

            # if NWR, then check gp, hov, rev
            if region == 'NWR':
                df_l = pd.Series(df_md.iloc[8:,0][df_md.iloc[8:,1]=='Y'])
                rev = df_l.apply(lambda x: 'R' in str(x))
                hov = df_l.apply(lambda x: 'H' in str(x))

                if sum(rev) > 1:
                    typ = 'rev'
                elif sum(hov) > 1:
                    typ = 'hov'
                else:
                    typ = 'gp'

            else:
                typ = 'gp'

            # List of sheets in commute file to get ngd
            shts = ['TT Summary', 'Jan Summary Stats', 'Feb Summary Stats', 
                    'Mar Summary Stats', 'Apr Summary Stats', 'May Summary Stats',
                   'Jun Summary Stats', 'Jul Summary Stats', 'Aug Summary Stats', 
                    'Sep Summary Stats', 'Oct Summary Stats', 'Nov Summary Stats', 
                    'Dec Summary Stats']

            ngd = list()

            # loop through and record ngd
            for sht in shts:
                try:
                    df_tt = xl_file.parse(sheet_name = sht)
                    ngd.append(df_tt.iloc[:,3].mean())
                except:
                    ngd.append(np.nan)

            xl_file.close()

            # create and append row to results dataframe
            row = {'Filename' : name, 'Length' : length, 'Type' : typ, 
                   'Tot GD' : ngd[0], 'Jan GD' : ngd[1], 'Feb GD' : ngd[2], 
                   'Mar GD' : ngd[3], 'Apr GD' : ngd[4], 'May GD' : ngd[5], 
                   'Jun GD' : ngd[6], 'Jul GD' : ngd[7], 'Aug GD' : ngd[8],
                   'Sep GD' : ngd[9], 'Oct GD' : ngd[10], 'Nov GD' : ngd[11], 
                   'Dec GD' : ngd[12]}

            trac_df = trac_df.append(row, ignore_index = True)

    # save results to .csv
    save_path = paths['com_xl_path']
    save_f = os.path.join(save_path,'%s commute_files.csv'%(region))

    trac_df.loc[:,['Filename', 'Length', 'Type', 'Tot GD', 
                   'Jan GD', 'Feb GD', 'Mar GD', 'Apr GD', 
                   'May GD', 'Jun GD', 'Jul GD', 'Aug GD', 
                   'Sep GD', 'Oct GD', 'Nov GD', 'Dec GD']].to_csv(save_f,
                                                                  index=False)

#------------------------------------------------------------------------------
def write_hist(region, objs, paths, variables):
    base_year = objs[objs.keys()[0]].base_year
    curr_year = objs[objs.keys()[0]].current_year

    if not isinstance(variables, list):
        variables = list(variables)

    p = paths['hist_cor_path']
    fname = '%s Historical.xlsx'%region
    f = os.path.join(paths['hist_cor_path'], fname)
    book = openpyxl.load_workbook(f)
    xl_i = pd.ExcelFile(f)
    xl_o = pd.ExcelWriter(f, engine='openpyxl')
    xl_o.book = book
    
    try:
        xl_o.save()
    except IOError:
        now = datetime.datetime.now()
        now = now.strftime("%Y%m%d-%H%M%S")   
        f = f.replace('.xlsx', ' %s.xlsx'%now)
        xl_o = pd.ExcelWriter(f, engine='openpyxl')
        xl_o.book = book    
    
    cors = [x.split()[0] for x in objs.keys() if x.endswith(region)]
    cors = sorted(map(int, cors))
    
    for v in variables:
        try:
            df = xl_i.parse(sheet_name=v.lower())
            del xl_o.book[v.lower()]
        except:
            df = pd.DataFrame(columns=cors + ['Total'])

        for b_c in ['base', 'current']:
            yr = getattr(objs[objs.keys()[0]], '%s_year'%b_c)

            if yr not in df.index:
                df.loc[yr, :] = np.nan

            for cor in cors:
                obj = objs['%s %s'%(cor, region)]
                df_sum = obj.summary
                idx = [i for i in df_sum.index if \
                       i.lower().startswith('%s %s'%(b_c, v))][0]
                df.loc[yr, int(cor)] = df_sum.loc[idx, 'Total']
                df = df.drop(columns=['Total'])
                df['Total'] = df.sum(axis=1)        
        
        df.to_excel(xl_o, sheet_name=v.lower())
    
    xl_o.save()
  
#------------------------------------------------------------------------------
def plot_hist(region, variables, paths):

    if not isinstance(variables, list):
        variables = list(variables)

    # Create directory structure if necessary
    s_path = os.path.join(paths['plot_path'], '10_Historical')
    if not os.path.exists(s_path):
        os.mkdir(s_path)
    
    font = {'ax' : {
            'family': 'Helvetica Neue LT Std',
            'weight': 'roman',
            'size': '6'},
            
            'title' : {
            'family': 'Helvetica Neue LT Std',
            'weight': 'medium',
            'size': '8'}
            }    
    
    matplotlib.rc('font', **font['ax'])

    w, h = [3.58, 1.25]
    lt, rt, bm, tp = [0.085, 0.99, 0.12, 0.8]    


    #
    fname = '%s Historical.xlsx'%region
    f = os.path.join(paths['hist_cor_path'], fname)
    xl = pd.ExcelFile(f)


    for v in variables:
        df = xl.parse(sheet_name=v.lower())
        if df.shape[0] > 10:
            df = df.iloc[-10:,:]
        
        for cor, col in df.iteritems():
            if v == 'vmt':
                col = col/1000.
                
            name = '%s %s'%(cor, region)
            M = col.max()
            n = len(str(int(M / 6))) - 1 
            intvl = int(round(M / 5,-n))
            intvl = intvl + (10**n) if intvl < M / 5 else intvl
            intvl = intvl if ((M / (intvl * 5.)) > 0.75) else \
                    intvl - (0.5 * 10**n)
            
            yticks = map(lambda x: x * intvl , list(range(0, 6)))

            # instantiate plot object
            fig, ax = plt.subplots(figsize = (w,h))
            fig.subplots_adjust(left=lt, right=rt, bottom=bm, top=tp)
            
            ax.plot(col, color = '0', lw = 0.75)

            # set gridlines, ticks, and invisible borders
            ax.grid(axis = 'y', lw = 0.25)
            ax.tick_params(axis = 'x', direction = 'in')
            ax.yaxis.set_tick_params(color = '1', pad=0)
            ax.spines['left'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.set_xticks(col.index)
            ax.set_yticks(yticks)
            
            ax.set_title(name, loc='left')

            fname = '%s historical %s.pdf'%(name, v)
            
            # save figure
            fig.savefig(os.path.join(s_path, fname), format = 'pdf')
            
            plt.close()
                            

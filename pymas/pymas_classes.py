#---------------------------- import dependencies ------------------------------
import numpy as np # 1.14.0
import pandas as pd # 0.22.0
import os
import openpyxl
import datetime
import pickle
import time
from pymas.corridor_tools import *

# Plotting
import matplotlib # '2.1.2'
from matplotlib.ticker import NullLocator
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt


#------------------------------------------------------------------------------
class Corridor:

    def __init__(self, route, region, paths,
                 base_year, curr_year, user=''):
        
        # Define attributes that are provided with arguments
        self.name = int(route)
        self.region = region
        self.base_year = int(base_year)
        self.current_year = int(curr_year)
        self.analyst = user
        
        # Filepaths
        self.paths = paths
        
        # Retrieve config files for corridor (results go to self.config_data)
        self._read_config()
        
        # Read in contour files from Trac (results go to self.contour_data)
        self._read_contour()
        
        # Initiate an empty dict for loop data
        self.loops = {}
        self.loop_list = get_loops(self.name, self.paths)
        
        # add vmt/pmt, delay, cost of delay, and GHG attributes
        self._calculate_vmt_pmt()
        self.calculate_delay(measure='delay', threshold=self.delay_threshold)
        self.calculate_delay(measure='cong', threshold=self.cong_threshold)
        self.calculate_cost(cost_of='delay', label='cd')
        self.calculate_cost(cost_of='cong', label='cc')
        self._calculate_ghg()
        self._summarize()
        
        return None
        
    #--------------------------------------------------------------------------
    def _read_config(self):

        
        # Open Statewide config file
        sw_f = os.path.join(self.paths['sw_cfg_path'], 'Statewide_config.xlsx')
        sw_xl = pd.ExcelFile(sw_f)
        
        # Read GHG and cost config sheets
        ghg = sw_xl.parse(sheet_name='GHG', 
                          header=[0,1], 
                          index_col=[0,1])
        
        cost = sw_xl.parse(sheet_name='Cost', 
                           header=[0,1])
        
        # Open corridor config file
        cor_f = os.path.join(self.paths['cor_cfg_path'], 
                             '%s_%s_config.xlsx'%(self.name,
                                                  self.region))
        cor_xl = pd.ExcelFile(cor_f)
        
        # Write config dataframes 
        for b_c in ['base', 'current']:
            
            yr = getattr(self, '%s_year'%(b_c))
            
            # set GHG_config
            ghg_cfg = ghg[yr]
            ghg_cfg.rename_axis(None, axis=1)
            ghg_cfg.reset_index(inplace=True)
            ghg_cfg.columns = ['min', 'max', 'ST', 'CT', 'Transit', 'Cars'] 

            setattr(self, '%s_ghg_config'%(b_c), ghg_cfg)
            
            # set cost config
            setattr(self, '%s_cost'%(b_c),
                    cost[yr].loc[self.region,:])    
            
            # set Corridor config
            setattr(self, '%s_config'%(b_c),
                    cor_xl.parse(sheet_name=str(yr)))        

        # Write plotting parameters
        try:
            df = cor_xl.parse(sheet_name='Plotting')
            pp = dict(keys=['heatmap', 'delay_axis', 'loopgroups'])
        except:
            pp = None
        
        if pp: 
            pp['heatmap'] = df.iloc[:, 0:3].dropna(how='all')
            pp['delay_axis'] = df.iloc[:, 3:6].dropna(how='all')
            pp['loopgroups'] = sorted(list(df['LoopGroups'].dropna()))
        
        self.plot_params = pp
        
        # Write miscellaneous variables
        for i, row in sw_xl.parse(sheet_name='Misc').iterrows():
            setattr(self, i, row[self.region])
        
        # Write peak times
        self.peak = sw_xl.parse(sheet_name='PeakTimes')[self.region]
        
        # Set corridor directions
        drct = sw_xl.parse(sheet_name='MP_directions')
        self.inc_dir = drct.loc[self.name, 'Increasing']
        self.dec_dir = drct.loc[self.name, 'Decreasing']
        
        return None
                
    #--------------------------------------------------------------------------
    def _read_contour(self):
        
        for b_c in ['base', 'current']:        
            for measure in ['Speed', 'Volume']:
                
                yr = getattr(self, '%s_year'%b_c)
                
                # list files in directory with [measure] and [corridor]
                xl = os.listdir(self.paths['%s_cont_path'%b_c])
                xl = [f for f in xl if \
                           measure in f and \
                           str(yr) in f and \
                           f.startswith('%s '%(self.name))]

                # if less or more than 2 contour files exist, raise exception
                if len(xl) != 2:
                    ex = ('There should be 2 files each (%s and %s) for %s %s %s'
                          %(self.inc_dir, self.dec_dir, self.name,
                          measure, yr))
                    
                    raise Exception(ex)

                # Read in contour file and assign as attribute
                for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
                    xl_d = [f for f in xl if ' %s '%(d.upper()) in f][0]
                    xl_d = os.path.join(self.paths['%s_cont_path'%b_c], xl_d)
                    df = pd.read_excel(xl_d, sheet_name='Fixed interval data')
                    
                    # Shift mileposts for SR 167
                    if self.name == 167:
                        df.columns = df.columns.map(lambda x: x - 1.5)
                    
                    # If float columns then round (ensures proper alignment)
                    if df.columns.dtype == float:
                        df.columns = df.columns.map(lambda x: round(x,1))
                    
                    cfg = self.base_config
                    cfg = cfg[cfg['SpdLmt_%s'%d.upper()] >0]
                    df = df.loc[:, cfg.index.values]
                    
                    setattr(self, '%s_%s_%s'%(b_c, measure.lower(), d),df)
    
        return None
    
    #---------------------------------------------------------------------------
    def _calculate_vmt_pmt(self):
        
        for b_c in ['base', 'current']:
            yr = getattr(self, '%s_year'%b_c)
            for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
                # assign variable names to config, length, volume, st% and ct%
                cfg = getattr(self, '%s_config'%b_c).copy()
                cfg = cfg[cfg['SpdLmt_%s'%d.upper()] > 0]
                l = cfg['Length']
                vol = getattr(self, '%s_volume_%s'%(b_c, d))
                st_pct = cfg['ST_pct'] # 'ST_pct_%s'%(direction)
                ct_pct = cfg['CT_pct'] # 'CT_pct_%s'%(direction)
                pass_pct = pd.Series(1., index = cfg.index) - (st_pct + ct_pct)
                
                # vmt = length (mi) * vol (veh/hr) / (12 epochs/hr)
                vmt = l * vol / 12 
                
                # assign to vmt attribute for total, passenger veh, and trucks
                setattr(self, '%s_vmt_%s'%(b_c, d), vmt)
                setattr(self, '%s_pass_vmt_%s'%(b_c, d), vmt * pass_pct)
                setattr(self, '%s_st_vmt_%s'%(b_c, d), vmt * st_pct)
                setattr(self, '%s_ct_vmt_%s'%(b_c, d), vmt * ct_pct)
                
                # pmt = vmt * occupancy
                am_occ = cfg['AM_Occ_%s'%d.upper()]
                pm_occ = cfg['PM_Occ_%s'%d.upper()]
                
                pmt = vmt.copy()
                pmt.iloc[0:144, :] =  pmt.iloc[0:144, :] * am_occ
                pmt.iloc[145:, :] = pmt.iloc[145:, :] * pm_occ
                
                setattr(self, '%s_pmt_%s'%(b_c, d), pmt)

        return None
    
    #---------------------------------------------------------------------------
    def calculate_delay(self, measure='delay', threshold=50./60.):
        
        # In case threshold was manually entered, overwrite the attribute
        setattr(self, '%s_threshold'%(measure), threshold)
        
        max_delay = []
        
        for b_c in ['base', 'current']:
            
            for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
                # set variables to config, vmt, and pmt
                cfg = getattr(self, '%s_config'%(b_c))
                cfg = cfg[cfg['SpdLmt_%s'%d.upper()] > 0]
                vmt = getattr(self, '%s_vmt_%s'%(b_c, d))
                p_vmt = getattr(self, '%s_pass_vmt_%s'%(b_c, d))
                st_vmt = getattr(self, '%s_st_vmt_%s'%(b_c, d))
                ct_vmt = getattr(self, '%s_ct_vmt_%s'%(b_c, d))
                pmt = getattr(self, '%s_pmt_%s'%(b_c, d))
                spd = getattr(self, '%s_speed_%s'%(b_c, d))
        
                # calculate delay threshold vector
                sl = cfg['SpdLmt_%s'%d.upper()]
                
                # total delay
                tot_delay = vect_delay(spd, vmt, sl, threshold)
                setattr(self, '%s_%s_%s'%(b_c, measure, d), 
                           tot_delay)
                
                max_delay.append(tot_delay.max().max())
                
                
                # passenger veh delay
                setattr(self, '%s_pass_%s_%s'%(b_c, measure, d),
                        vect_delay(spd, p_vmt, sl, threshold))
                
                # single trucks delay
                setattr(self, '%s_st_%s_%s'%(b_c, measure, d),
                        vect_delay(spd, st_vmt, sl, threshold))
                    
                # combination trucks delay
                setattr(self, '%s_ct_%s_%s'%(b_c, measure, d), 
                        vect_delay(spd, ct_vmt, sl, threshold))
                     
                # person delay
                setattr(self, '%s_pers_%s_%s'%(b_c, measure, d),
                        vect_delay(spd, pmt, sl, threshold))
        
            # record maximum delay
            if measure == 'delay':
                self.max_delay = max(max_delay)
        
        return None
    
    #---------------------------------------------------------------------------
    def calculate_cost(self, cost_of='delay', label='cd'):

        '''Before running this method, calculate_delay needs to have been run
        with the desired measure and threshold. 
        
        E.g., to calculate cost of delay with respect to speed limit: 
        self.calculate_delay(measure='delSpdLmt', threshold=1.0)
        self.calculate_cost(cost_of='delSpdLmt', name='cdsl')
        '''
        
        for b_c in ['base', 'current']:
            yr = getattr(self, '%s_year'%b_c)
            for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
                
                # set variables to delay and cost
                pv_del = getattr(self, '%s_pass_%s_%s'%(b_c, cost_of, d))
                pers_del = getattr(self, '%s_pers_%s_%s'%(b_c, cost_of, d))
                st_del = getattr(self, '%s_st_%s_%s'%(b_c, cost_of, d))
                ct_del = getattr(self, '%s_ct_%s_%s'%(b_c, cost_of, d))
                
                cost = getattr(self, '%s_cost'%b_c)
                pv_cost = cost['PV']
                pers_cost = cost['PP']
                st_cost = cost['LT'] # LT ~ ST
                ct_cost = cost['HT'] # HT ~ CT
                
                # assuming occupancy of 1 (vmt = pmt) for trucks
                # cost of delay = s * $/veh
                st_cd = st_cost * st_del
                ct_cd = ct_cost * ct_del
                
                # passenger vehicle cost of delay = veh-hrs * $/veh
                pv_cd = pv_cost * pv_del
                
                # passenger vehicle passenger cost of delay:
                # person hours = total person hours - truck person hours
                p_del = pers_del - (st_del + ct_del)
                p_cd = pers_cost * p_del
                
                # sum cost of delay and write to attribute
                setattr(self, '%s_%s_%s'%(b_c, label, d),
                           st_cd + ct_cd + pv_cd + p_cd)
                
        return None
                
    #---------------------------------------------------------------------------
    def _calculate_ghg(self):
        
        for b_c in ['base', 'current']:
            yr = getattr(self, '%s_year'%b_c)
            for d in [self.inc_dir, self.dec_dir]:
                
                # set variables to config, ghg_config, speed, and vmt
                cfg = getattr(self, '%s_config'%(b_c))
                ghg_cfg = getattr(self, '%s_ghg_config'%(b_c))
                spd = getattr(self, '%s_speed_%s'%(b_c, d.lower()))
                p_vmt = getattr(self, '%s_pass_vmt_%s'%(b_c, d.lower()))
                st_vmt = getattr(self, '%s_st_vmt_%s'%(b_c, d.lower()))
                ct_vmt = getattr(self, '%s_ct_vmt_%s'%(b_c, d.lower()))
                
                # set initial dataframes for ghg factors as speed copies
                p_f = spd.copy() # passenger vehicle factors dataframe
                st_f = spd.copy() # single truck factors dataframe
                ct_f = spd.copy() # combination truck factors dataframe
                
                # loop through ghg_config and replace speeds with factors
                for i, row in ghg_cfg.iterrows():
                    # assign low and high values for speed range
                    low = row['min']
                    high = row['max']
                    
                    # replace values within range with factors
                    p_f[(p_f >= low) & (p_f < high)] = row['Cars']
                    st_f[(st_f >= low) & (st_f < high)] = row['ST']
                    ct_f[(ct_f >= low) & (ct_f < high)] = row['CT']
                    
                # ghg = conv (lb/g) * factor (g/mi) * vmt (mi)
                conv = 0.00220462262 #lb/g
                
                p_ghg = conv * p_f * p_vmt
                st_ghg = conv * st_f * st_vmt
                ct_ghg = conv * ct_f * ct_vmt
                
                # sum GHG and assign to attribute
                setattr(self, '%s_ghg_%s'%(b_c, d.lower()), 
                        p_ghg + ct_ghg + st_ghg)
            
        return None
    
    #---------------------------------------------------------------------------
    def _summarize(self):
        
        # assign variable names for directions
        d1 = self.inc_dir.lower()
        d2 = self.dec_dir.lower()
        
        # List measures to summarize
        measures = ([x[4:-1] for x in self.__dict__.keys() if \
                    x.endswith(d1) and \
                    x.startswith('base')]) # Have corresponding current & d2
    
        for m in measures:
            # get total for directions (i.e. N + S)
            for b_c in ['base', 'current']:
                yr = getattr(self, '%s_year'%b_c)
                # assign variables to dataframes
                m_d1 = getattr(self, '%s%s%s'%(b_c, m, d1))
                m_d2 = getattr(self, '%s%s%s'%(b_c, m, d2))
                
                # calculate sum of directions
                m_tot = m_d1.fillna(0) + m_d2.fillna(0)
                if m_tot.sum().sum() == 0:
                    m_tot = pd.merge(m_d1, m_d2, how='outer',
                                     left_index=True, right_index=True)
                setattr(self, '%s%stotal'%(b_c, m), m_tot)
                
            
            # get delta and pct change across years 
            for d in [d1, d2, 'total']:
                m_base = getattr(self, 'base%s%s'%(m, d))
                m_curr = getattr(self, 'current%s%s'%(m, d))
                
                setattr(self, 'delta%s%s'%(m, d), m_curr - m_base)
                setattr(self, 'pct%s%s'%(m, d), 
                        vect_pct_change(m_base, m_curr))
        
        # create delay dataframe
        df_cols = ['Base Delay %s'%(d1.upper()), 
                   'Current Delay %s'%(d1.upper()), 
                   'Pct Change Delay %s'%(d1.upper()),
                   'Base Delay %s'%(d2.upper()), 
                   'Current Delay %s'%(d2.upper()), 
                   'Pct Change Delay %s'%(d2.upper()),
                   'Base Delay Total',
                   'Current Delay Total',
                   'Pct Change Delay Total']
        
        
        del_df = pd.DataFrame(columns = df_cols,
                             index = self.base_config.index)
        
        
        # write to results dataframe
        del_df[df_cols[0]] = getattr(self, 'base_delay_%s'%(d1)).sum()
        del_df[df_cols[1]] = getattr(self, 'current_delay_%s'%(d1)).sum()
        del_df[df_cols[2]] = vect_pct_change(del_df[df_cols[0]],
                                            del_df[df_cols[1]])
        
        del_df[df_cols[3]] = getattr(self, 'base_delay_%s'%(d2)).sum()
        del_df[df_cols[4]] = getattr(self, 'current_delay_%s'%(d2)).sum()
        del_df[df_cols[5]] = vect_pct_change(del_df[df_cols[3]],
                                            del_df[df_cols[4]])
        
        del_df[df_cols[6]] = getattr(self, 'base_delay_total').sum()
        del_df[df_cols[7]] = getattr(self, 'current_delay_total').sum()
        del_df[df_cols[8]] = vect_pct_change(del_df[df_cols[6]],
                                            del_df[df_cols[7]])
        
        setattr(self, 'delay_comparison', del_df)
        
        # Create Summary dataframe
        summary_rows = ['Base Delay (veh-hr/day)', 
                        'Current Delay (veh-hr/day)',
                        'Base Delay (veh-hr/year)', 
                        'Current Delay (veh-hr/year)',
                        'Change in Delay (%)', 
                        'Base GHG Emissions (lb CO2 eq./year)',
                        'Current GHG Emissions (lb CO2 eq./year)',
                        'Change in GHG Emissions (%)',
                        'Base Cost of Delay (dollars/year)',
                        'Current Cost of Delay (dollars/year)',
                        'Change in Cost of Delay (%)',
                        'Base vmt (/day)',
                        'Current vmt (/day)',
                        'Base vmt (/year)',
                        'Current vmt (/year)',
                        'Change in vmt (%)',
                        'Base pmt (/year)',
                        'Current pmt (/year)',
                        'Change in pmt (%)',
                        'Current Truck %',
                        'Region',
                        'Corridor',
                        'Extent',
                        'Base Year',
                        'Current Year',
                        'Date of Analysis',
                        'Analyst'
                       ]
        
        summary_cols = [d1.upper(), d2.upper(), 'Total']
        
        summary_df = pd.DataFrame(index = summary_rows, columns = summary_cols)
        
        # step through rows in summary df
        row = 0
        for var in ['delay_day', 'delay_yr', 'ghg_yr',
                    'cd_yr', 'vmt_day', 'vmt_yr', 'pmt_yr']:
            col = 0
            v, dur = var.split('_')
            n = 1 if dur == 'day' else self.ndays
            nrows = 2 if dur == 'day' else 3
            
            for d in [d1, d2, 'total']:
                b = getattr(self, 'base_%s_%s'%(v, d)).sum().sum() * n
                c = getattr(self, 'current_%s_%s'%(v, d)).sum().sum() * n
                
                summary_df.iloc[row, col] = b
                summary_df.iloc[row + 1, col] = c
                
                if nrows == 3:
                    ch = pct_change(b, c)
                    summary_df.iloc[row + 2, col] = ch
                
                col+=1

            row += nrows
        
        # calculate truck percentage
        tp = ((self.current_st_vmt_total.sum().sum() +
              self.current_ct_vmt_total.sum().sum()) /
              self.current_vmt_total.sum().sum())        
        
        summary_df.loc['Current Truck %', d1.upper()] = tp
        
        # add metadata
        summary_df.loc['Region', d1.upper()] = self.region
        summary_df.loc['Corridor', d1.upper()] = self.name
        summary_df.loc['Extent', d1.upper()] = '%s-%s'%(
                                                  self.base_config.index.min(),
                                                  self.base_config.index.max())
        summary_df.loc['Base Year', d1.upper()] = self.base_year
        summary_df.loc['Current Year', d1.upper()] = self.current_year
        summary_df.loc['Date of Analysis', d1.upper()] = datetime.datetime.now()
        summary_df.loc['Analyst', d1.upper()] = self.analyst
        
        setattr(self, 'summary', summary_df)        
        
        return None
        
    #---------------------------------------------------------------------------
    def add_loops(self, lgs=None):
        
        if lgs:
            loopgroups = lgs
        else:
            if not self.loop_list:
                self.loop_list = get_loops(self.name, self.paths)
            loopgroups = self.loop_list
        
        if not isinstance(loopgroups, list):
            loopgroups = [loopgroups]
        
        for lg in loopgroups:
            if int(lg[0:3]) != self.name:
                print('%s is not on corridor %s'%(lg, self.name))
                continue
            
            fname = '%s.dat'%lg
            fpath = p = os.path.join(self.paths['base_loop_path'],'..',fname)
            
            try:
                obj = read_object(lg, 'LoopGroup', self.paths)
            except IOError:
                print('Building %s LoopGroup object'%lg)
                obj = LoopGroup(lg,
                                self.base_year, 
                                self.current_year, 
                                self.paths)
            
            if obj.base_ngd or obj.current_ngd:
                self.loops[lg] = obj 
        
        self._summarize_lg_data()
        
        return None
    
    #---------------------------------------------------------------------------
    def export_excel(self, variables, suffix=''):
        
        xl_f = os.path.join(self.paths['cor_out_path'], 
                            '%s %s%s.xlsx'%(self.name,
                                            self.region,
                                            suffix))
        
        try:        
            wrt = pd.ExcelWriter(xl_f, datetime_format='h:mm:ss')

            sheets = ['Summary', 'Delay_Comparison']

            for v in variables:
                for yr in ['Pct', 'Delta', 'Base', 'Current']:
                    for d in [self.inc_dir, self.dec_dir]:
                        sheets.append('%s_%s_%s'%(yr, v, d))

            for sheet in sheets:
                self.__dict__[sheet.lower()].to_excel(wrt, sheet_name=sheet)


            wrt.save()
        
        except IOError:
            print(xl_f + ' could not be written')
    
        return None
    
    #---------------------------------------------------------------------------
    def export_dat(self, suffix=''):
        
        dat_f = os.path.join(self.paths['cor_out_path'],
                             '%s %s%s.dat'%(self.name,
                                            self.region,
                                            suffix))
        
        try:
            with open(dat_f, 'wb') as f:
                pickle.dump(self, f)
        
        except IOError:
            print(dat_f + ' could not be written')
    
        return None
    
    #---------------------------------------------------------------------------
    def plot_heatmap(self, variables, norm):

        #If necessary, make variables a list
        if not isinstance(variables, list):
            variables = [variables]
        
        
        #If necessary, create subfolder in plot directory
        hm_path = os.path.join(self.paths['plot_path'], '2_Heatmaps')
        if not os.path.exists(hm_path):
            os.mkdir(hm_path)
        
        for v in variables:
            for b_c in ['pct', 'delta', 'base', 'current']:
                for d in [self.inc_dir, self.dec_dir]:            

                    # read in dataframe, transpose and sort
                    df = getattr(self, '%s_%s_%s'%(b_c, v, d.lower()))
                    df = df.T.sort_index(ascending=False)
                    
                    # if a norm value is passed in use it else max 
                    if v == 'delay':
                        v_max = norm if norm else df.max().max()
                    else:
                        v_max = df.max().max()
                        
                    # create figure
                    fig, ax = plt.subplots(figsize = (4,4))

                    # plot heatmap
                    if df.min().min() < 0: 
                        ax.imshow((abs(df[df < 0])), 
                           cmap='Blues', # blue = negative
                           vmin=0, vmax=v_max,
                           aspect='auto') 

                        ax.imshow((abs(df[df > 0])), 
                           cmap = 'Reds', # red = positive
                           vmin = 0, vmax = v_max,
                           aspect = 'auto') 

                    else:
                        ax.imshow((v_max - df), 
                               cmap='gray', # map to greyscale
                               vmin=0, vmax=v_max, 
                               aspect='auto') 

                    # turn off the axes and frame
                    ax.axis('off')

                    # set locator so the plot does not include whitespace
                    ax.xaxis.set_major_locator(NullLocator())
                    ax.yaxis.set_major_locator(NullLocator())

                    # save to pdf and close
                    f_name = '%s %s %s %s_%s.pdf'%(self.name, 
                                                   d, 
                                                   self.region, 
                                                   v, 
                                                   b_c)
                    
                    # Define and create subfolder (if necessary)
                    f_path = os.path.join(hm_path, v)
                    
                    if not os.path.exists(f_path):
                        os.mkdir(f_path)

                    f_path = os.path.join(f_path, f_name)

                    try:
                        pdffig = PdfPages(f_path)
                    except IOError:
                        now = datetime.datetime.now()
                        now = now.strftime("%Y%m%d-%H%M%S")
                        f_path = f_path.replace('.pdf', ' %s.pdf'%now)
                        pdffig = PdfPages(f_path)

                    fig.savefig(pdffig, format = 'pdf',
                                bbox_inches = 'tight', 
                                pad_inches = -0.05) # remove margins

                    metadata = pdffig.infodict()
                    metadata['Title'] = '%s %s %s %s Heatmap %s'\
                                            %(self.name,
                                            d,
                                            self.region,
                                            v,
                                            b_c)
                    metadata['Author'] = self.analyst + '(Generated by PyMAS)'
                    metadata['Subject'] = 'Heatmap of %s'%(v) + \
                                          'by milepost and time of day'

                    pdffig.close()                    
                    plt.close('all')
                    
        return None

    #---------------------------------------------------------------------------
    def plot_heatmaps(self, variables, norm=None):

        if self.plot_params is None:
            print('plot_heatmaps requires plotting parameters to be ' +
                  'defined in the %s_%s_config.xlsx file.'%(self.name,
                                                            self.region))
            return None
        
        # Define required variables
        dirs = {'N' : 'Northbound',
                'E' : 'Eastbound',
                'S' : 'Southbound',
                'W' : 'Westbound'}
        ds = [self.dec_dir, self.inc_dir]
        arwx = [.425, .575]
        arws = ['-|>', '<|-']
        b_am, e_am, b_pm, e_pm = self.peak.tolist()
        
        llim, ulim = min(self.base_config.index), max(self.base_config.index)
        
        # set plot parameters
        font = {'ax' : {
                'family': 'Helvetica Neue LT Std',
                'weight': 'roman',
                'size': '6'},

                'title' : {
                'family': 'Helvetica Neue LT Std',
                'weight': 'medium',
                'size': '8'}
                }    
        
        #If necessary, make variables a list
        if not isinstance(variables, list):
            variables = [variables]
        
        #If necessary, create subfolder in plot directory
        hm_path = os.path.join(self.paths['plot_path'], '2_Heatmaps')
        if not os.path.exists(hm_path):
            os.mkdir(hm_path)    

        for v in variables:
            for b_c in ['delta', 'base', 'current', 'current2']:
                
                if b_c.endswith('2'):
                    b_c = b_c[0:-1]
                    i = 1
                    sec = '_main'
                else:
                    i = 0
                    sec = '_appx'
                
                # Set first height for appendix, second for main report
                w = 5.5
                h = self.plot_params['heatmap'].loc[i,'Height'] # plot area
                h_f = h + .5 #height of full figure in inches
                bot = 0.15 / h_f
                top = (h_f - 0.25) / h_f
                mp_lab = top + .25 * (0.25 / h_f)
                
                # Instantiate plot object
                fig, axes = plt.subplots(nrows=1, ncols=2, figsize=(w, h_f))
                axes[0].yaxis.tick_right()
                fig.subplots_adjust(left=0.01, right=0.99, bottom=bot, 
                                    top=top, wspace=0.7, hspace=0)

                # add labels
                fig.text(0.38, mp_lab, 'MP', ha='left', fontdict=font['ax'])
                fig.text(0.62, mp_lab, 'MP', ha='right', fontdict=font['ax'])
                for i, row in self.plot_params['heatmap'].iterrows():
                    mp = row['MP']
                    loc = bot + ((top - bot) * (mp - llim) / (ulim - llim)) 
                    fig.text(0.38, loc, '%g'%mp, 
                             ha='left', va='center', fontdict=font['ax'])
                    fig.text(0.5, loc, row['Label'], 
                             ha='center', va='center', fontdict=font['ax'])
                    fig.text(0.62, loc, '%g'%mp, 
                             ha='right', va='center', fontdict=font['ax'])

                for i in [0,1]:
                    ax = axes[i]
                    d = ds[i]
                    ax.set_title(dirs[d], fontdict=font['title'])
                     
                    # read in dataframe, transpose and sort
                    df = getattr(self, '%s_%s_%s'%(b_c, v, d.lower()))
                    df = df.T.sort_index(ascending=True)
                    
                    # if a norm value is passed in use it else max 
                    if v == 'delay':
                        v_max = norm if norm else df.max().max()
                    else:
                        v_max = df.max().max()

                    kwargs = {'vmin' : 0, 'vmax' : v_max,
                              'aspect' : 'auto', 'origin' : 'lower',
                              'extent' : [0,287, llim, ulim]}
                    
                    # plot heatmap
                    if df.min().min() < 0: 
                        ax.imshow((abs(df[df < 0])), 
                           cmap='Blues', # are negative
                           **kwargs) 

                        ax.imshow((abs(df[df > 0])), 
                           cmap = 'Reds', # are positive
                           **kwargs) 

                    else:
                        ax.imshow((v_max - df), 
                               cmap='gray', # map to greyscale
                               **kwargs) 
                        
                    ax.set_frame_on(False)
                    ax.tick_params(axis='both', which='both', bottom=False, 
                                   top=False, left=False, right=False, pad=0)
                    pad_am = '  ' if e_am - b_am < 5 else ''
                    pad_pm = '  ' if e_pm - b_pm < 5 else ''
                    ax.set_xticklabels(['%s AM%s'%(b_am, pad_am),
                                        '%s%s AM'%(pad_am, e_am), 
                                        '%s PM%s'%(b_pm-12, pad_pm), 
                                        '%s%s PM'%(pad_pm, e_pm-12)], 
                                       fontdict=font['ax'])
                    ax.set_xticks([b_am*12, e_am*12, b_pm*12, e_pm*12])
                    ax.set_yticks([])
                    for val in self.plot_params['heatmap'].MP.tolist():
                        ax.axhline(val, lw=0.1, color='0.5', alpha=0.5)
                    am = matplotlib.patches.Rectangle((b_am*12, llim), 
                                                      (e_am-b_am)*12, 
                                                      ulim, 
                                                      facecolor='0', alpha=0.02)
                    pm = matplotlib.patches.Rectangle((b_pm*12, llim), 
                                                      (e_pm-b_pm)*12, 
                                                      ulim, 
                                                      facecolor='0', alpha=0.02)
                    ax.add_patch(am)
                    ax.add_patch(pm)
                    ax.annotate('', xy=(arwx[i], bot), xytext=(arwx[i], top), 
                                xycoords='figure fraction', 
                                arrowprops={'facecolor':'black', 
                                            'arrowstyle' : arws[i], 
                                            'lw' : 0.2})

                # save to pdf and close
                f_name = '%s %s %s_%s%s.pdf'%(self.name, 
                                              self.region, 
                                              v, 
                                              b_c,
                                              sec)

                # Define and create subfolder (if necessary)
                f_path = os.path.join(hm_path, v)

                if not os.path.exists(f_path):
                    os.mkdir(f_path)

                f_path = os.path.join(f_path, f_name)

                try:
                    pdffig = PdfPages(f_path)
                except IOError:
                    now = datetime.datetime.now()
                    now = now.strftime("%Y%m%d-%H%M%S")
                    f_path = f_path.replace('.pdf', ' %s.pdf'%now)
                    pdffig = PdfPages(f_path)

                fig.savefig(pdffig, format = 'pdf')

                metadata = pdffig.infodict()
                metadata['Title'] = '%s %s %s Heatmap %s'\
                                        %(self.name,
                                        self.region,
                                        v,
                                        b_c)
                metadata['Author'] = self.analyst + '(Generated by PyMAS)'
                metadata['Subject'] = 'Heatmap of %s'%(v) + \
                                      'by milepost and time of day'

                pdffig.close()                    
                plt.close('all')
                
        return None              
                                 
    #---------------------------------------------------------------------------
    def plot_line(self, variables):   
        # set plot parameters
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
        w, h = [3.58, 1.1]
        lt, rt, bm, tp = [0.085, 0.97, 0.12, 0.8]

        
        #If necessary, create subfolder in plot directory
        line_path = os.path.join(self.paths['plot_path'], '1_CorridorLines')
        if not os.path.exists(line_path):
            os.mkdir(line_path)
        
        
        if not isinstance(variables, list):
            variables = [variables]
        
        for v in variables:
            for d in [self.inc_dir, self.dec_dir, 'total']:
                fig, ax = plt.subplots(figsize = (w,h))
                fig.subplots_adjust(left=lt, right=rt, bottom=bm, top=tp)                
                
                # Calculate Series of mean of speed or sum of other variables
                if v.lower() == 'speed':
                    b_srs = getattr(self, 'base_%s_%s'%(v, d.lower())).mean()
                    c_srs = getattr(self, 'current_%s_%s'%(v, d.lower())).mean()
                
                else:
                    b_srs = getattr(self, 'base_%s_%s'%(v, d.lower())).sum()
                    c_srs = getattr(self, 'current_%s_%s'%(v, d.lower())).sum()
                
                # if index is not float, reindex to keep order and set labels
                if b_srs.index.dtype != float:
                    ax.set_xticklabels(b_srs.index.values)
                    
                    b_srs.reset_index(drop=True, inplace=True)
                    c_srs.reset_index(drop=True, inplace=True)
                else:
                    ax.set_xlim(left = b_srs.index.min(), 
                                right = b_srs.index.max())
                    
                ax.plot(b_srs, 
                        color = '0.5', lw = 0.75) # Gray is always base year
                ax.plot(c_srs, 
                        color = '0', lw = 0.75) # Black is always current year

                ax.set_title('%s %s %s; %s and %s'%(self.name,
                                              d,
                                              v,
                                              self.base_year,
                                              self.current_year),
                             loc='left')

                # if variable is delay set y axis to those in config
                # otherwise use default
                if v.lower() == 'delay' and self.plot_params:
                    # set y axis parameters
                    yticks = list(self.plot_params['delay_axis']['Delay_%s'%d])
                    ax.set_yticks(yticks)
                    ax.set_ylim(bottom=0, top=max(yticks))
                
                ax.grid(axis = 'y', lw = 0.25)
                ax.tick_params(axis='x', direction='in')
                ax.yaxis.set_tick_params(color='1', pad=0)


                # remove border
                for side in ['left', 'right', 'top']:
                    ax.spines[side].set_visible(False)


                # save figure and add metadata
                # save to pdf and close
                f_name = '%s %s %s %s.pdf'%(self.name, 
                                            d, 
                                            self.region, 
                                            v)

                # Define and create (if necessary) subfolder for variable
                f_path = os.path.join(line_path, v)

                if not os.path.exists(f_path):
                    os.mkdir(f_path)

                f_path = os.path.join(f_path, f_name)                
                
                try:
                    pdffig = PdfPages(f_path)
                except IOError:
                    now = datetime.datetime.now()
                    now = now.strftime("%Y%m%d-%H%M%S")
                    f_path = f_path.replace('.pdf', ' %s.pdf'%now)
                    pdffig = PdfPages(f_path)
                    
                fig.savefig(pdffig, format = 'pdf')

                metadata = pdffig.infodict()
                metadata['Title'] = ('%s %s %s %s Lineplot %s-%s'
                                     %(self.name, d, self.region, 
                                     v, self.base_year, 
                                     self.current_year))

                metadata['Author'] = self.analyst + '(Generated by PyMAS)'
                metadata['Subject'] = ('Lineplot of %s by milepost %s-%s'
                                       %(v,self.base_year, 
                                         self.current_year))

                pdffig.close()
                plt.close('all')
                    
        return None
    
    #---------------------------------------------------------------------------
    def _summarize_lg_data(self):
        
        for d in [self.inc_dir, self.dec_dir]:   
                      
            loops = [l for l in self.loops.keys() if l[12] == d]
        
            if len(loops) == 0:
                setattr(self, 'loop_summary_%s'%d.lower(), None)
                continue
            
            res_df = pd.DataFrame()
            
            for b_c in ['base', 'current']:

                b_c_df = pd.DataFrame(index = loops, columns = ['PctGD', 
                                      'Delay_agg', 'Delay_day', 
                                      'Speed', 'Volume'])     

                for l in loops:
                    b_c_df.loc[l,'PctGD'] = getattr(self.loops[l], 
                                                  '%s_ngd'%b_c)/261.
                    b_c_df.loc[l,'Delay_agg'] = getattr(self.loops[l], 
                                                        '%s_del_agg'%b_c)
                    b_c_df.loc[l,'Delay_day'] = getattr(self.loops[l], 
                                                        '%s_del_day'%b_c)
                    b_c_df.loc[l,'Speed'] = getattr(self.loops[l], 
                                                    '%s_spd'%b_c)
                    b_c_df.loc[l,'Volume'] = getattr(self.loops[l], 
                                                     '%s_vol'%b_c)
                
                # index by milepost
                b_c_df.index = b_c_df.index.map(lambda x: float(x[5:10])/100)
                b_c_df = b_c_df.sort_index()
                b_c_df.columns = b_c_df.columns.map(lambda x: '%s_%s'%(b_c, x))
                
                res_df = pd.concat([res_df, b_c_df], axis=1)
            
            setattr(self, 'loop_summary_%s'%d.lower(), res_df)
                
        return None
    
    #---------------------------------------------------------------------------
    def plot_lg_data(self, by_day_fmt='png', by_mp_fmt='pdf'):
               
        try:
            w = self.base_config.index.max() - self.base_config.index.min()
        except:
            return None
        
        for nm, obj in sorted(self.loops.iteritems()):
            obj.plot_by_day(by_day_fmt)
        
        for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
            fig = plt.figure(figsize=(w, 8), dpi=150)

            # create axes
            ax1 = plt.subplot2grid((50,2),(1,0), colspan=2, rowspan=5)
            ax2 = plt.subplot2grid((50,2),(6,0), colspan=2, rowspan=25)
            ax3 = plt.subplot2grid((50,2),(31,0), colspan=2, rowspan=10)
            ax4 = plt.subplot2grid((50,2),(41,0), colspan=2, rowspan=10)

            # set x limits to min and max mileposts and set tick parameters
            for ax in [ax1, ax2, ax3, ax4]:
                ax.tick_params(
                   axis='both',       # changes apply to both axes
                   which='both',      # both major and minor ticks are affected
                   top=False,         # ticks along the top edge are off
                   labelbottom=False) # labels along the bottom edge are off

            # add label back to bottom frame
            ax4.tick_params(labelbottom=True)

            df = self.__dict__['loop_summary_%s'%d]
            
            for b_c in ['base', 'current']:
                
                color = '0' if b_c == 'current' else '0.5'
                
                try:    
                    # If corridor has loop summary data, plot lines
                    
                    
                    df['%s_PctGD'%b_c].plot(ax=ax1, color=color, lw=0.5)
                    df['%s_Delay_agg'%b_c].plot(ax=ax2, color=color, 
                                                ls='--', lw=0.5)
                    df['%s_Delay_day'%b_c].plot(ax=ax2, color=color, lw=0.5)
                    df['%s_Speed'%b_c].plot(ax = ax3, color=color, lw=0.5)
                    df['%s_Volume'%b_c].plot(ax=ax4, color=color, lw=0.5)
                
                    # Plot points for where nearby points are missing
                    df_m = df[(df['%s_PctGD'%b_c].shift() == 0) & \
                              (df['%s_PctGD'%b_c].shift(-1) == 0)]
                    
                    if df_m.shape[0] != 0:
                        df_m['%s_Delay_agg'%b_c].plot(ax=ax2, color=color, 
                                                      lw=0, marker='_')
                        df_m['%s_Delay_day'%b_c].plot(ax=ax2, color=color, 
                                                      lw=0, marker='.')
                        df_m['%s_Speed'%b_c].plot(ax=ax3, color=color, 
                                                  lw=0, marker='.')
                        df_m['%s_Volume'%b_c].plot(ax=ax4, color=color, 
                                                   lw=0, marker='.')
                
                except Error:
                    print('%s is missing data for %s'%(self.name, 
                                                       getattr(self, 
                                                               '%s_year'%b_c)))

            # Set figure title, axis labels
            fig.suptitle('%s %s Delay, Speed, Volume by MP'%(self.name, 
                                                             d.upper()))
            
            ax1.set_ylabel('% Data')
            ax2.set_ylabel('Delay (veh-hr/day/mile)')
            ax3.set_ylabel('Speed (mph)')
            ax4.set_ylabel('Volume (veh/day)')
            
            # Create directory structure if necessary
            f_path = os.path.join(self.paths['plot_path'], '9_Loops')
            if not os.path.exists(f_path):
                os.mkdir(f_path)
            
            f_path = os.path.join(f_path, 'byMP')
            if not os.path.exists(f_path):
                os.mkdir(f_path)
            
            f_name = '%s %s DSV_MP.%s'%(self.name, d.upper(), by_mp_fmt)
            f_path = os.path.join(f_path, f_name)

            # Save figure
            try:
                fig.savefig(f_path, format=by_mp_fmt, bbox='tight')
        
            except IOError:
                now = datetime.datetime.now()
                now = now.strftime("%Y%m%d-%H%M%S")
                f_path = f_path.replace('.%s'%by_mp_fmt, 
                                        ' %s.%s'%(now, by_mp_fmt))
                fig.savefig(f_path, format=by_mp_fmt, bbox='tight')
            
            plt.close('all') 

        return None
            
    #---------------------------------------------------------------------------  
    def plot_throughput(self, lgs=None, PctGD=0.4, fmt='png'):
        
        # if loopgroups not provided, default to lgs with greater than PctGD of
        # good days -OR- if 'config', get loopgroups from 
        if not lgs:
            lgs = self._throughput_lgs(PctGD=PctGD)
        elif lgs == 'config':
            try:
                lgs = self.plot_params['loopgroups']
            except:
                return None
        
        for lg in lgs:
            try:
                self.loops[lg].plot_throughput(fmt=fmt)
            except KeyError:
                print('%s not in %s %s Corridor object'%(lg,
                                                         self.name,
                                                         self.region))
                
        return None
    
    #---------------------------------------------------------------------------  
    def update_plot_params(self):
        
        # Open corridor config file
        cor_f = os.path.join(self.paths['cor_cfg_path'], 
                             '%s_%s_config.xlsx'%(self.name,
                                                  self.region))
        cor_xl = pd.ExcelFile(cor_f)
        
        # Write plotting parameters
        try:
            df = cor_xl.parse(sheet_name='Plotting')
            pp = dict(keys=['heatmap', 'delay_axis', 'loopgroups'])
        except:
            pp = None
        
        if pp: 
            pp['heatmap'] = df.iloc[:, 0:3].dropna(how='all')
            pp['delay_axis'] = df.iloc[:, 3:6].dropna(how='all')
            pp['loopgroups'] = sorted(list(df['LoopGroups'].dropna()))
        
        self.plot_params = pp       
        
    #---------------------------------------------------------------------------
    def _throughput_lgs(self, PctGD=0.4):
        
        lgs = []
        for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
            df = getattr(self, 'loop_summary_%s'%d)
            
            if isinstance(df, pd.DataFrame):
                mps = df[(df.base_PctGD > PctGD) & \
                         (df.current_PctGD > PctGD)].index
                mps = mps.map(lambda x: '%03des%05d_M%s__'%(self.name, 
                                                            round(x * 100, 1), 
                                                            d.upper()))
                lgs = lgs +  list(mps) 
        
        return lgs

    #---------------------------------------------------------------------------
    def export_excel_lg(self, lgs=None, suffix=''):
               
        # if loopgroups not provided default to get loopgroups from plot_params
        if not lgs:
            try:
                lgs = self.plot_params['loopgroups']
            except:
                return None
        
        # prepare Excel writer
        xl_f = '%s %s_thru_loops.xlsx'%(self.name, self.region)
        xl_p = self.paths['cor_out_path']
        xl_f = os.path.join(self.paths['cor_out_path'], 
                            '%s %s_thru_loops%s.xlsx'%(self.name,
                                                       self.region,
                                                       suffix))

        try:        
            wrt = pd.ExcelWriter(xl_f, datetime_format='h:mm:ss')
            for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
                df = getattr(self, 'loop_summary_%s'%d)
                df.to_excel(wrt, sheet_name='loop_summary_%s'%d)
            wrt.save()

        except IOError:
            now = datetime.datetime.now()
            now = now.strftime("%Y%m%d-%H%M%S")
            xl_f = xl_f.replace('.xlsx', ' %s.xlsx'%now)
            wrt = pd.ExcelWriter(xl_f, datetime_format='h:mm:ss')
            for d in [self.inc_dir.lower(), self.dec_dir.lower()]:
                df = getattr(self, 'loop_summary_%s'%d)
                df.to_excel(wrt, sheet_name='loop_summary_%s'%d)

        # loop through loop summaries and combine years into one dataframe
        for lg in sorted(lgs):
            obj = self.loops[lg]

            try:
                df1 = obj.base_by_time.loc[:, ['Speed', 'Volume', 'Thru']]
                df1.columns = df1.columns.map(lambda x: 'base_%s'%x)
            except AttributeError:
                df1 = pd.DataFrame()

            try:
                df2 = obj.current_by_time.loc[:, ['Speed', 'Volume', 'Thru']]
                df2.columns = df2.columns.map(lambda x: 'current_%s'%x)
            except AttributeError:
                df2 = pd.DataFrame()


            df = pd.merge(df1, df2, how='outer', 
                          left_index=True, right_index=True)

            if df.shape[1] == 6:
                df = df[df.iloc[:, 2] + df.iloc[:, 5] < 2]
            else:
                df[df.iloc[:, 2] < 1]

            # write to Excel
            df.to_excel(wrt, sheet_name=lg)

        wrt.save()

        return None
    
        
#------------------------------------------------------------------------------
class LoopGroup:
    
    def __init__(self, loop_name, base_year, curr_year, paths, user=''):
        
        self.name = loop_name
        self.corridor = int(loop_name.split(' ')[0][0:3])
        self.direction = loop_name.split('_')[1][1]
        self.mp = float(loop_name[5:10])/100.
        self.base_year = int(base_year)
        self.current_year = int(curr_year)
        self.analyst = user        
        
        # Filepaths
        self.paths = paths

        # Populate object from Excel file
        self.process_lg_xl()
        
        # Calculate throughput
        self.calc_throughput()
        
        return None
    
    
    #-------------------------------------------------------------------------- 
    def process_lg_xl(self):
        
        for b_c in ['base', 'current']:
            
            yr = getattr(self, '%s_year'%b_c)
            
            # Define file name, path, and Excel object
            lp_f = '%s M-F %s.xlsx'%(self.name, yr)
            
            lp_p = self.paths['%s_loop_path'%b_c]
            
            
            # if empty file, set attributes to None
            try:
                lp_xl = pd.ExcelFile(os.path.join(lp_p, lp_f))
            except:
                setattr(self, '%s_ngd'%(b_c), 0.0)
                for attr in ['by_day', 'by_time', 'nlanes', 'max_vplph',
                            'del_agg', 'del_day', 'spd', 'vol']:
                    setattr(self, '%s_%s'%(b_c, attr), None)
                continue
                

            # Read Speed and Volume Sheets
            df_spd = lp_xl.parse(sheet_name='Speed')
            df_spd[df_spd.applymap(lambda x: x < 0)] = np.nan

            df_vol = lp_xl.parse(sheet_name='Volume')
            df_vol[df_vol.applymap(lambda x: x < 0)] = np.nan
            
            # calculate delay
            df_del = vect_delay(df_spd, df_vol, 60, 0.83333)

            # weight speed by volume
            df_spd = (df_spd * df_vol).sum() / df_vol.sum() #vector result
            spd = (df_spd * df_vol.sum()).sum() / df_vol.sum().sum() #scalar
            
            # create dataframe with speed, volume, delay, and % data by day
            df_d = pd.DataFrame([df_spd, df_vol.sum(), 
                                 df_del.sum(), df_del.count()/288.],
                                 index = ['Speed', 'Volume', 'Delay', 'Pct']).T

            # set to attribute
            setattr(self, '%s_by_day'%b_c, df_d)
            
            
            # read in summary stats dataframe
            df_SS = lp_xl.parse(sheet_name='Summary Stats')
            df_SS = df_SS.set_index('Time')
            
            # extract number of lanes
            nl = (12 * df_SS.iloc[0,1]) / df_SS.iloc[0,2]
            setattr(self, '%s_nlanes'%b_c, nl)

            # get number of good days
            cols = [c for c in df_SS.columns if self.name[0:10] in c \
                    and df_SS.loc[:, c].sum() > 0]
            ngd = df_SS.loc[:, cols].mean().mean()
            setattr(self, '%s_ngd'%b_c, ngd)
            
            # get maximum VpLpH
            setattr(self, '%s_max_vplph'%b_c, max(df_SS['Mean VpLpH']))

            # calculate delay from summary (annual averages)
            df_SS['Delay_agg'] = vect_delay(df_SS['Mean Speed'],
                                        df_SS['Mean Volume'],
                                        60, 0.8333333)
            # delay from daily sheet
            df_SS['Delay_day'] = df_del.sum(axis=1)/ngd
            
            # set to attribute
            df_SS = df_SS.rename(columns={'Mean Speed' : 'Speed', 
                                       'Mean Volume' : 'Volume'})
            setattr(self, '%s_by_time'%b_c,
               df_SS.iloc[:-3,:].loc[:,['Speed', 
                                        'Volume', 
                                        'Delay_agg',
                                        'Delay_day']])
            
            setattr(self, '%s_spd'%b_c, spd)
            setattr(self, '%s_vol'%b_c, df_SS['Volume'].sum())
            setattr(self, '%s_del_agg'%b_c, df_SS['Delay_agg'].sum())
            setattr(self, '%s_del_day'%b_c, df_SS['Delay_day'].sum())
            
            
        return self
    
    #--------------------------------------------------------------------------   
    def calc_throughput(self, thr=50):
        
        M_list = [self.base_max_vplph, self.current_max_vplph]
        M = max(M_list)
        
        if M:
            for b_c in ['base', 'current']:
                df_SS = getattr(self, '%s_by_time'%b_c)
                try:
                    cong = df_SS.Speed < thr
                except:         
                    continue
                M_v = getattr(self, '%s_nlanes'%b_c) * M / 12
                
                df_SS.loc[:, 'Thru'] = 1.
                
                LC = df_SS[cong].Volume.map(lambda x: M_v - x) / M_v
                df_SS.loc[cong, 'Thru'] = 1. - LC

                setattr(self, '%s_by_time'%b_c, df_SS)
                
        return None
            
    #-------------------------------------------------------------------------- 
    def plot_throughput(self, fmt='pdf'):
        if self.base_ngd > 0 and self.current_ngd > 0:
            #If necessary, create subfolder in plot directory
            thru_path = os.path.join(self.paths['plot_path'], '5_Throughput')
            if not os.path.exists(thru_path):
                os.mkdir(thru_path)        

            # Initiate plot
            w, h = [3.58, 1.1]
            lt, rt, bm, tp = [0.085, 0.97, 0.12, 0.8]
            fig, ax = plt.subplots(figsize=(w, h), dpi=150)
            fig.subplots_adjust(left=lt, right=rt, bottom=bm, top=tp)
            
            # Set plot parameters
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

            # set gridlines, ticks, and invisible borders
            ax.grid(axis = 'y', lw = 0.25)
            ax.tick_params(axis = 'x', direction = 'in')
            ax.yaxis.set_tick_params(color = '1', pad=0)
            ax.spines['left'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.set_xticks([5*12, 8*12, 11*12, 14*12, 17*12, 20*12])
            ax.set_xticklabels(['5 AM', '8 AM', '11 AM', 
                                '2 PM', '5 PM', '8 PM'])
            ax.set_xlim(left = 5*12, right = 20*12)

            # Set percentage y axis
            ax.set_yticks([1,0.8,0.6,0.4,0.2])
            ax.set_yticklabels(['100%', '80%', '60%', '40%', '20%'])
            ax.set_ylim(bottom = 0, top = 1.01)      

            # Plot
            ax.plot(self.base_by_time.Thru.values, color='0.5', lw=0.75)
            ax.plot(self.current_by_time.Thru.values, color='0', lw=0.75)

            # Set title
            vphpl = max(self.base_max_vplph, self.current_max_vplph)
            ax.set_title('%s : %s vphpl'%(self.name, 
                                          int(vphpl)), 
                         loc='left')

            # save figure
            f_path = os.path.join(thru_path, '%s.%s'%(self.name, fmt))
            
            if fmt == 'pdf':
                try:
                    pdffig = PdfPages(f_path)

                except IOError:
                    now = datetime.datetime.now()
                    now = now.strftime("%Y%m%d-%H%M%S")
                    f_path = f_path.replace('.pdf', ' %s.pdf'%now)
                    pdffig = PdfPages(f_path)

                fig.savefig(pdffig, format='pdf')
            
                # write pdf metadata
                base_year = self.base_year
                curr_year = self.current_year

                metadata = pdffig.infodict()
                metadata['Title'] = '%s Throughput %s-%s'%(self.name, 
                                                       base_year, curr_year)
                metadata['Author'] = self.analyst + '(Generated by PyMAS)'
                sub = 'Throughput productivity %s & %s'%(base_year, curr_year)
                metadata['Subject'] = sub

                pdffig.close()
            
            fig.savefig(f_path, format=fmt)
            plt.close()       

        return None
        
    #--------------------------------------------------------------------------  
    def plot_by_day(self, fmt='pdf'):
        
        fig = plt.figure()

        # create axes
        ax1 = plt.subplot2grid((50,2),(1,0), colspan=2, rowspan = 5)
        ax2 = plt.subplot2grid((50,2),(6,0), colspan=2, rowspan = 25)
        ax3 = plt.subplot2grid((50,2),(31,0), colspan=2, rowspan = 10)
        ax4 = plt.subplot2grid((50,2),(41,0), colspan=2, rowspan = 10)
        

        for b_c in ['base', 'current']:

            try:
                df = self.__dict__['%s_by_day'%b_c]
                df['Pct'].plot(ax=ax1, color='0')
                df['Delay'].plot(ax=ax2, color='0')
                df['Speed'].fillna(0).plot(ax=ax3, color='0')
                df['Volume'].plot(ax=ax4, color='0')
                
            except TypeError:
                print('%s is missing data for %s'%(self.name, 
                                                   getattr(self, '%s_year'%b_c)))

        # set x limits to min and max mileposts and set tick parameters
        for ax in [ax1, ax2, ax3, ax4]:
            ax.tick_params(
               axis='both',          # changes apply to the x-axis
               which='both',      # both major and minor ticks are affected
               bottom=False,      # ticks along the bottom edge are off
               top=False,         # ticks along the top edge are off
               #left=False,
               right=False,
               #labelleft=False,
               labelbottom=False) # labels along the bottom edge are off
            ax.set_xlim(datetime.datetime(self.base_year, 1, 1),
                        datetime.datetime(self.current_year, 12, 31))
        
        ax1.set_ylim(-.1,1.1)
        ax1.set_yticks([0,1])
        ax1.set_yticklabels(['0%','100%'])
        ax2.set_ylim(0,2200)
        ax2.set_yticks([0,500,1000,1500,2000])
        ax3.set_ylim(30,70)
        ax3.set_yticks([45,60])
        ax4.set_ylim(0,200000)
        ax4.set_yticks([0,50000,100000,200000])
        ax4.tick_params(labelbottom=True)
        
        fig.suptitle('%s Delay, Speed, Volume by Day'%(self.name))
        fig.text(0.5, 0.8, '% Data', horizontalalignment='center')
        fig.text(0.5, 0.6, 'Delay (veh-hr/mile)', horizontalalignment='center')
        fig.text(0.5, 0.33, 'Speed (mph)', horizontalalignment='center')
        fig.text(0.5, 0.2, 'Volume (veh)', horizontalalignment='center')
        
        
        # Create directory structure if necessary
        f_path = os.path.join(self.paths['plot_path'], '9_Loops')
        if not os.path.exists(f_path):
            os.mkdir(f_path)

        f_path = os.path.join(f_path, 'byday')
        if not os.path.exists(f_path):
            os.mkdir(f_path)

        f_path = os.path.join(f_path, '%s%s'%(self.corridor, self.direction))
        if not os.path.exists(f_path):
            os.mkdir(f_path)            
        
   
        f_name = '%s DSV_day.%s'%(self.name, fmt)
        f_path = os.path.join(f_path, f_name)
        
        # Save figure
        try:
            fig.savefig(f_path, format=fmt)
        except IOError:
            now = datetime.datetime.now()
            now = now.strftime("%Y%m%d-%H%M%S")            
            f_path = f_path.replace('.%s'%fmt, ' %s.%s'%(now, fmt))
            fig.savefig(f_path, format=fmt)
        
        plt.close('all')       

        return None
    
#------------------------------ Region class ----------------------------------
class Region:
    
    def __init__(self, com_xl, paths, user=''):
        
        self.com_xl = com_xl
        self.paths = paths
        self.analyst = user
        self.name = com_xl.split('_')[0]
        
        # Read
        self.read_com_xl()
        
        # Add commutes
        self.commutes = {}
        self.add_commutes()
        self.summarize()
    
    #--------------------------------------------------------------------------  
    def read_com_xl(self):
        df = pd.read_excel(os.path.join(self.paths['com_xl_path'], 
                                        self.com_xl), 
                         sheet_name = 'commutes', header = 1, index_col = 0)
        
        df = df[[x for x in df.columns if 'Unnamed' not in x]]
        
        df['include'] = df.index.map(lambda x: str(x) == 'nan')
        df.reset_index(drop = True, inplace = True)
        df.index = df.Commute
        include = df['include'].iloc[1:]
        df = df.drop(columns = ['Commute', 'include'])
        
        com_paths = df.iloc[0, 1:]
        repl = pd.Series(index=com_paths.index)
        repl[0:len(repl)/2] = self.paths['base_com_path']
        repl[len(repl)/2:] = self.paths['current_com_path']
        com_paths[com_paths.isna()] = repl

        df = df.iloc[1:,:]

        self.base_year = int(com_paths.index.values[0].split('_')[1])
        self.current_year = int(com_paths.index.values[-1].split('_')[1])
        
        self.files = df
        self.include = include
        self.com_paths = com_paths
        
        return None
    
    #--------------------------------------------------------------------------  
    def add_commutes(self, com_list=None):
        
        if com_list is None:
            coms = self.include[self.include == True].index
        else:
            coms = com_list
        
        for com in coms:
            print('Building %s Commute object'%com)
            com_obj = Commute(com, self.name, self.files.loc[com, :],
                              self.com_paths, self.paths, self.analyst)
            self.commutes[com] = com_obj
        
        return None

    #--------------------------------------------------------------------------  
    def update(self):
        self.read_com_xl()
        self.add_commutes()
        self.summarize()
    
    #--------------------------------------------------------------------------  
    def summarize(self):
        # GP
        for b_c in ['base', 'current']:
            try:
                gp_df = getattr(self, '%s_gp_summary'%b_c)
                hov_df = getattr(self, '%s_hov_summary'%b_c)
            except AttributeError:
                xl = os.path.join(self.paths['tmp_path'], 
                                    'Commute_Template.xlsx')
                cols = pd.read_excel(xl, sheet_name='columns')
                gp_cols = cols['GP'].iloc[1:]
                hov_cols = cols['HOV'].iloc[1:]
                gp_df = pd.DataFrame(index=self.files.index, 
                                     columns=gp_cols)
                hov_df = pd.DataFrame(index=self.files.index, 
                                      columns=hov_cols)
        
            for com, obj in self.commutes.iteritems():
                # fill in gp_df
                gp_df.loc[com, 'dir'] = obj.direction
                gp_df.loc[com, 'len'] = getattr(obj, '%s_gp_distance'%b_c)
                gp_df.loc[com, 'ttsl'] = getattr(obj, '%s_gp_ttsl'%b_c)
                gp_df.loc[com, 'mt3'] = getattr(obj, '%s_gp_mt3'%b_c)
                gp_df.loc[com, '80pct'] = getattr(obj, '%s_gp_80pct'%b_c)
                
                obj_sum = getattr(obj, '%s_gp_summary'%b_c)
                for per in ['am', 'pm']:
                    for col in obj_sum.columns:
                        gp_df.loc[com, '%s_%s'%(per, col)] = \
                            obj_sum.loc[per, col]  
                
                n_hovs = 0
                if obj.hov:        
                    n_hovs += 1
                    # fill in hov_df
                    gp = ['gphov'] if obj.gphov else ['gp']
                    hov = ['hov', 'rev'] if obj.rev else ['hov'] 

                    # direction
                    hov_df.loc[com, 'dir'] = obj.direction

                    for gphov in gp + hov:
                        # gphov variable for writing dataframe 
                        v = 'gp' if gphov==gp[0] else gphov
                        i = 0 if gphov == gp[0] else 1
                        cols = ['time', 'avg_tt', '95pct'][i:]
                        obj_sum = getattr(obj, '%s_%s_summary'%(b_c, gphov))
                        obj_sum = obj_sum.loc[:, cols]

                        # Route information
                        hov_df.loc[com, '%s_len'%v] = \
                            getattr(obj, '%s_%s_distance'%(b_c, gphov))
                        hov_df.loc[com, '%s_ttsl'%v] = \
                            getattr(obj, '%s_%s_ttsl'%(b_c, gphov))
                        hov_df.loc[com, '%s_mt3'%v] = \
                            getattr(obj, '%s_%s_mt3'%(b_c, gphov))

                        for per in ['am', 'pm']:
                            for col in obj_sum.columns:
                                hov_df.loc[com, '%s_%s_%s'%(per, v, col)] = \
                                    obj_sum.loc[per, col]
                    
            setattr(self, '%s_gp_summary'%b_c, gp_df)
            
            if n_hovs:
                setattr(self, '%s_hov_summary'%b_c, hov_df)
                self.hov = True
            else:
                self.hov = False
            
        return None

    #---------------------------------------------------------------------------
    def export_excel(self, suffix=''):
            
        xl_tmp = os.path.join(self.paths['tmp_path'], 
                            'Commute_Template.xlsx')
        book = openpyxl.load_workbook(xl_tmp)
        
        for b_c in ['base', 'current']:
            xl_out = '%s_commutes_%s%s.xlsx'%(self.name, b_c, suffix)
            xl_out = os.path.join(self.paths['com_out_path'], xl_out)
            
            try:
                writer = pd.ExcelWriter(xl_out, engine='openpyxl')
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                writer.save()
            except IOError:
                now = datetime.datetime.now()
                now = now.strftime("%Y%m%d-%H%M%S")            
                xl_out = xl_out.replace('.xlsx', ' %s.xlsx'%now)                
                writer = pd.ExcelWriter(xl_out, engine='openpyxl')
            
            writer.book = book
            writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
            
            
            gphovs = ['GP', 'HOV'] if self.hov else ['GP']
            
            for gphov in gphovs:
                df = getattr(self, '%s_%s_summary'%(b_c, gphov.lower())).copy()
                df = df.reset_index()
                df.to_excel(writer, sheet_name=gphov, startrow=3, 
                            header=False, index=False)
                df_footer = pd.DataFrame(
                    [['Analysis Year:', getattr(self, '%s_year'%b_c)],
                     ['Date of Analysis:', datetime.datetime.now()],
                     ['Analyst:', self.analyst]])
                if gphov == 'GP':
                    ri = df['80pct'].sum() / df['mt3'].sum()
                    ri_row = pd.DataFrame([['Reliability Index:', ri]])
                    df_footer = df_footer.append(ri_row)
                                                    
                df_footer.to_excel(writer, sheet_name=gphov, startrow=30,
                                   header=False, index=False)
                
            writer.save()
    
    #---------------------------------------------------------------------------
    def export_dat(self, suffix=''):

        dat_f = os.path.join(self.paths['com_out_path'],
                             '%s_commutes%s.dat'%(self.name,
                                                  suffix))

        try:
            with open(dat_f, 'wb') as f:
                pickle.dump(self, f)
        
        except IOError:
            print(dat_f + ' could not be written')
    
        return None
    
    #---------------------------------------------------------------------------
    def plot(self, plot_types=['hov_cong', 'gp_cong', 'gp_sc', 
                               'gp_tt', 'gp_spd']):
                               
        '''
        Acceptable plot_types are: 
        hov_cong : plots same year congestion - HOV vs GP vs REV
        gp_cong : plots comparison of base and current year congestion
        gp_sc : plots comparison of base and current year severe congestion 
        gp_tt : plots comparison of base and current year travel time
        gp_spd : plots comparison of base and current year speed
        '''
        
        for nm, obj in self.commutes.iteritems():
            if nm not in self.include[self.include == True].index:
                continue
            
            if obj.hov:
                obj.plot(plot_types)
            else:
                ptypes = [pt for pt in plot_types if pt.startswith('gp_')]
                obj.plot(ptypes)
    
    
#------------------------------ Commute class ---------------------------------
class Commute:
    
    def __init__(self, commute, region, files, com_paths, paths, user=''):
        
        # Filepaths
        self.paths = paths
        self.com_paths = com_paths
        
        # Define attributes that can be retreived from files dataframe
        self.name = commute
        self.direction = files[0]
        self.files = files.iloc[1:]
        self.base_year = int(files.index[1].split('_')[1])
        self.current_year = int(files.index[-1].split('_')[1])
        self.region = region
        self.analyst = user
        
        # Read in data and subset vmt from Corridor object
        self._read_commutes()
        self._read_config()
        self._get_extents()
        self.subset_variables(['vmt'])
        self.gp_analysis()
        if self.hov: 
            self.hov_analysis()
        
        return None
        
    #---------------------------------------------------------------------------
    def _read_commutes(self):
        
        base = self.files[0:len(self.files)/2].index.values
        
        for idx, fname in self.files.iteritems():
            
            b_c = 'base' if idx in base else 'current'
                       
            gphov = idx.split('_')[0].lower()
            gphov = ''.join([x for x in gphov if x.isalnum()])
            
            if str(fname) == 'nan':
                setattr(self, '%s_%s'%(b_c, gphov), False)
                continue
            else:
                setattr(self, '%s_%s'%(b_c, gphov), True)
            
            # Read Excel data
            xl = pd.ExcelFile(os.path.join(self.com_paths[idx], fname))
            
            # Travel Time data
            df_tt = pd.read_excel(xl, sheet_name='TT Summary')
            setattr(self, '%s_%s_tt'%(b_c, gphov), df_tt)
            
            # Metadata
            df_m = pd.read_excel(xl, sheet_name='Trip Metadata', header = None)
            setattr(self, '%s_%s_md'%(b_c, gphov), df_m)
            setattr(self, '%s_%s_distance'%(b_c, gphov), float(df_m.iloc[1,1]))
        
        for gphov in ['gphov', 'hov', 'rev']:
            try:
                b = getattr(self, 'base_%s'%gphov)
                c = getattr(self, 'current_%s'%gphov)
                setattr(self, gphov, bool(b + c))
                delattr(self, 'base_%s'%gphov)
                delattr(self, 'current_%s'%gphov)
            except AttributeError:
                setattr(self, gphov, False)
        
        return None
                
    #---------------------------------------------------------------------------
    def _read_config(self):                

        xl_f = os.path.join(self.paths['sw_cfg_path'], 'Statewide_config.xlsx')
        xl_obj = pd.ExcelFile(xl_f)
        
        # directions
        self.corr_id = xl_obj.parse(sheet_name='MP_directions')          

        # cost
        cost = xl_obj.parse(sheet_name='Cost',  header=[0,1])
        for b_c in ['base', 'current']:
            yr = getattr(self, '%s_year'%b_c)
            setattr(self, '%s_cost'%b_c,
                    cost[yr].loc[self.region,:])            
        
        # variables from "Misc" sheet
        for var, val in xl_obj.parse(sheet_name='Misc').iterrows():
            setattr(self, var, val[0])
        
        # Peak hours and corresponding indices
        self.peak = xl_obj.parse(sheet_name='PeakTimes')[self.region]
        self.ampeak = list(range(int(self.peak.b_am * 12), 
                                 int(self.peak.e_am * 12)))
        self.pmpeak = list(range(int(self.peak.b_pm * 12), 
                                 int(self.peak.e_pm * 12)))   
        
        return None
    #---------------------------------------------------------------------------
    def _get_extents(self):
        
        # loop through metadata dataframes
        for md in [x for x in self.__dict__.keys() if x.endswith('_md')]:
            com = md[:-3]
            
            df = getattr(self, md)
            
            if df.iloc[3,0] == 'Corridor': # ER, SCR
                # Create extents dataframe
                setattr(self, com + '_ext',
                    pd.DataFrame({'Corridor' : [int(df.iloc[3,1])],
                                 'min' : [np.nan],
                                 'max' : [np.nan],
                                 'Direction' : [df.iloc[4,1]],
                                 'Inc_Dec' : [np.nan]}))
            
            else: # Get extents from loop data
                # Create series of loopgroups
                df = pd.Series(df.iloc[8:,0][df.iloc[8:,1]=='Y'])

                setattr(self, com + '_loops', df)
                
                # Extract corridor and milepost from loopgroup
                df = df.apply(lambda x: x.split('_')[0].split('es'))
                df = df.apply(lambda x: pd.Series(x, 
                                                  index = ['Corridor', 'mp']))
                df['Corridor'] = df['Corridor'].map(lambda x: int(x))
                df['mp'] = pd.to_numeric(df['mp'])/100. 
                idx_order = df['Corridor'].unique()
                
                # Groupby corridor and get min , max, and starting milepost
                df = df.groupby(by = 'Corridor')\
                       .agg(['min', 'max', lambda x : x.iloc[0]])
                df = df.loc[idx_order,:]
                df.reset_index(inplace = True)        

                # Name columns
                df.columns = ['Corridor', 'min', 'max', 'Direction']

                # returns 1 if start = max and 0 if start = min
                df['Direction'] = 1*(df['Direction'] == df['min'])
                df['Inc_Dec'] = df['Direction'].apply(lambda x: \
                                    'Increasing' if x==1 else 'Decreasing')
                
                # lookup direction (N, E, S, or W) from corr_id dataframe 
                df['Direction'] = df[['Corridor', 'Inc_Dec']]\
                                    .apply(lambda (x,y): self.corr_id.loc[x, y], 
                                           axis = 1)

                # Create extents dataframe
                setattr(self,com + '_ext', df)

        self.corridors = list(self.base_gp_ext['Corridor'])
        
        return None
                
    #---------------------------------------------------------------------------
    def subset_variables(self, variables):    

        if not isinstance(variables, list):
            variables = [variables] 
        
        # Create dictionary of corridor objects
        cor_objs = {}
        for cor in self.corridors:
            #get corridor object
            obj_f = os.path.join(self.paths['cor_out_path'], 
                                 '%s %s.dat'%(cor, self.region))
            with open(obj_f, 'r') as f:
                cor_objs[str(cor)] = pickle.load(f)
        
        # loop through gp_ext dataframes
        for b_c in ['base', 'current']:
            # get year_GPHOV and year (base or curr) from ext
            com = b_c + '_gp'
            ext = com + '_ext'
            
            # Create results dataframes 
            for v in variables:
                setattr(self, '%s_gp_%s'%(b_c, v),
                    pd.DataFrame(index = range(288)))
            
            cfg_res = pd.DataFrame()
            
            # Loop through rows of extents dataframe and subset data
            for i, row in getattr(self, ext).iterrows():
                
                cor = row['Corridor']
                m = row['min']
                M = row['max']
                d = row['Direction']
                asc = True if row['Inc_Dec'] == 'Increasing' else False
                
                cor_obj = cor_objs[str(cor)]
                
                for v in variables:
                    
                    try:
                        df = getattr(cor_obj, '%s_%s_%s'%(b_c, v, d.lower()))
                    except:
                        continue
                    
                    if str(m) == 'nan':
                        mps = df.columns
                        srt = False
                    else:
                        # if no data within extents, seek nearest
                        mps = []
                        i = 0
                        while len(mps) == 0:
                            mps = [x for x in df.columns if \
                                   x > m - i and x < M + i]
                            i +=0.01 
                        srt = True
                    
                    df = df[mps]
                    df.reset_index(inplace=True, drop=True)  
                    
                    # sort df ascending
                    if srt:
                        df.sort_index(axis=1, ascending=asc, inplace=True)
                    
                    # add corridor as prefix to column names
                    df.columns = df.columns.map(
                                     lambda x : '%s_%s'%(cor, x))
                    
                    # merge df with results dataframe
                    res = getattr(self, '%s_gp_%s'%(b_c, v))
                    res = res.merge(df, how='outer', left_index=True, 
                                    right_index=True)
                    setattr(self, '%s_%s'%(com, v), res)
                    
                # get corridor config
                cor_cfg = getattr(cor_obj, '%s_config'%b_c).loc\
                             [mps, ['Length', 'SpdLmt_%s'%d, 'AM_Occ_%s'%d, 
                                    'PM_Occ_%s'%d]]
                cor_cfg.sort_index(axis=0, ascending=asc, inplace=True)
                cor_cfg.index = cor_cfg.index.map(lambda x : '%s_%s'%(cor, x))
                
                # set speed limit
                cfg_res = pd.concat([cfg_res, cor_cfg])
                #setattr(self, '%s_cor_cfg'%com, cor_cfg)
            
            # collapse spdlmt and occupancy columns
            for col in ['SpdLmt', 'AM_Occ', 'PM_Occ']:
                cols = [c for c in cfg_res.columns if c.startswith(col)]
                cfg_res.loc[:, col] = cfg_res.loc[:, cols].sum(skipna=True, 
                                                               axis=1)
            
            setattr(self, '%s_cor_cfg'%com, cfg_res.loc[:,['Length', 
                                              'SpdLmt', 'AM_Occ', 'PM_Occ']])

            # calculate weighted average of spdlmt and occupancy
            for var in ['SpdLmt', 'AM_Occ', 'PM_Occ']:
                v = ((cfg_res[var] * cfg_res['Length']).sum() / 
                      cfg_res['Length'].sum())
                setattr(self, '%s_%s'%(com, var.lower()), v)
            
            # calculate travel time at sl and max throughput speed
            gps = ['gp', 'gphov'] if self.gphov else ['gp']
            
            for gp in gps:
                sl = getattr(self, '%s_gp_spdlmt'%b_c)
                ttsl = 60 * getattr(self, '%s_%s_distance'%(b_c, gp)) / sl 
                mt3 = ttsl / self.mt_threshold
                ttcong = ttsl / self.cong_threshold

                setattr(self, '%s_%s_ttsl'%(b_c, gp), ttsl)
                setattr(self, '%s_%s_mt3'%(b_c, gp), mt3)
                setattr(self, '%s_%s_ttcong'%(b_c, gp), ttcong)
            
        return None
            
    #---------------------------------------------------------------------------
    def gp_analysis(self):    
        
        # Get gp travel time attribute names
        atts = [x for x in self.__dict__.keys() if \
                'gp' in x and x.endswith('_tt')]   
        
        for att in atts:
            df = getattr(self, att).copy()
            com = att[:-3]
            df = df.rename(columns={'Median Speed' : '50 %ile'}) # for ER
            df = df.loc[:,['Time', 'Avg. TTS', '50 %ile', '80 %ile', 
                           '90 %ile', '95 %ile']]
            df.columns = ['time', 'avg_tt', '50pct', '80pct', '90pct', '95pct']
            df.iloc[:,1:] = df.iloc[:,1:].apply(lambda x: x/60.)
            
            # 80 percentile from 5a-8p for Results WA
            pct80 = df.loc[list(range(5*12, 20*12)), '80pct'].mean()
            setattr(self, '%s'%att.replace('_tt', '_80pct'), pct80)
            
            df = df[(df['avg_tt'] == max(df.loc[self.ampeak, 'avg_tt'])) | \
                    (df['avg_tt'] == max(df.loc[self.pmpeak, 'avg_tt']))]
            df.index = ['am', 'pm']
            df.loc[:, 'mt3i'] = df.avg_tt / getattr(self, '%s_mt3'%com)
            
            if '_gp_' in att:
                cong = self._congestion_analysis(att)
                df = pd.concat([df, cong], axis=1)
                
                vmt = self._sum_vmt(att)
                df = pd.concat([df, vmt], axis=1)
                
            setattr(self, '%s_summary'%com, df)
        
        return None
               
    #---------------------------------------------------------------------------
    def hov_analysis(self):
       
        if not self.hov:
            return None
        
        gphov = 'gphov' if self.gphov else 'gp'
        hovs = ['hov', 'rev'] if self.rev else ['hov']
        
        for b_c in ['base', 'current']:
            for hov in hovs:
                gp_df = getattr(self, '%s_%s_summary'%(b_c, gphov))

                hov_df = getattr(self, '%s_%s_tt'%(b_c, hov)).copy()
                hov_df = hov_df.loc[:,['Time', 'Avg. TTS', '95 %ile']]
                hov_df.columns = ['time', 'avg_tt', '95pct']
                hov_df.iloc[:,1:] = hov_df.iloc[:,1:].apply(lambda x: x/60.)
                
                hov_df = hov_df[(hov_df['time'] == gp_df.loc['am', 'time']) | \
                                (hov_df['time'] == gp_df.loc['pm', 'time'])]
                hov_df.index = ['am', 'pm']
                
                l = getattr(self, '%s_%s_distance'%(b_c, hov))
                sl = getattr(self, '%s_gp_spdlmt'%b_c)

                # set attributes
                setattr(self, '%s_%s_ttsl'%(b_c, hov), 
                        60 * l / sl)
                setattr(self, '%s_%s_mt3'%(b_c, hov), 
                        60 * l / (sl * self.mt_threshold))
                
                setattr(self, '%s_%s_summary'%(b_c, hov), hov_df)
                
        return None 
                
    #---------------------------------------------------------------------------
    def _congestion_analysis(self, tt_att):        
        
        b_c = tt_att.split('_')[0]
        df = getattr(self, tt_att).copy()
        sl = getattr(self, tt_att.replace('_tt', '_spdlmt'))
        th = self.cong_threshold
        ttcong = getattr(self, tt_att.replace('_tt', '_ttcong'))

        ampm = {'am' : list(range(0,144)), 'pm' : list(range(144,288))}

        df = df[df['Avg. Speed'] < sl*th]
        df.loc[:, 'exc_tt'] = (df.loc[:, 'Avg. TTS'] / 60) - ttcong

        res_df = pd.DataFrame(columns = ['sc', 'dc', 'cc'], 
                              index = ['am', 'pm'])

        for pk, idx in ampm.iteritems():       
            l = [x for x in df.index.values if x in idx]
            if l:
                # Start and duration
                st = df.loc[l[0], 'Time']
                dur = datetime.time(int((5 * len(l))/60), 
                                     (5 * len(l)) % 60)

                # Cost
                vmt = getattr(self, '%s_gp_vmt'%b_c).loc[l, :].sum(axis=1)
                exc_tt = df.loc[l, 'exc_tt'] / 60
                occ = getattr(self, '%s_gp_%s_occ'%(b_c, pk))
                c_v_h = self.base_cost['PV'] + occ * self.base_cost['PP']

                cost = (exc_tt * c_v_h * vmt).sum() / vmt.sum()

            else:
                st = np.nan
                dur = np.nan
                cost = np.nan

            res_df.loc[pk, 'sc'] = st
            res_df.loc[pk, 'dc'] = dur
            res_df.loc[pk, 'cc'] = cost
        
        return res_df
    
    #---------------------------------------------------------------------------
    def _sum_vmt(self, tt_att):        
        
        res_df = pd.DataFrame(columns = ['vmt'], 
                              index = ['am', 'pm'])        
        df = getattr(self, tt_att.replace('_tt', '_vmt'))
        ext = getattr(self, tt_att.replace('_tt', '_ext')) 
        cfg = getattr(self, tt_att.replace('_tt', '_cor_cfg'))
        
        for pk in ['am', 'pm']:
            idx = getattr(self, '%speak'%pk)

            vmt = 0
            for i, row in ext.iterrows():
                cols = [c for c in df.columns if \
                        c.startswith('%s_'%(row.Corridor))]
                len_df = cfg.loc[cols, 'Length'].sum()         
                
                if str(row['min']) == 'nan':
                    crct = 1.
                else:
                    len_act = row['max'] - row['min']
                    crct = len_act / len_df
                
                vmt += df.loc[idx, cols].sum().sum() * crct
                
            res_df.loc[pk, 'vmt'] = vmt

        return res_df
    
    #---------------------------------------------------------------------------
    def plot(self, plot_types=['hov_cong', 'gp_cong', 'gp_sc', 
                               'gp_tt', 'gp_spd']):
        '''
        Acceptable plot_types are: 
        hov_cong : plots same year congestion - HOV vs GP vs REV
        gp_cong : plots comparison of base and current year congestion
        gp_sc : plots comparison of base and current year severe congestion 
        gp_tt : plots comparison of base and current year travel time
        gp_spd : plots comparison of base and current year speed
        '''
        
        if not isinstance(plot_types, list):
            plot_types = [plot_types] 
        
        for ptype in plot_types:       

            # get GP/HOV from plot_type and define a boolean variable for HOV
            GPHOV, measure = ptype.split('_')
            HOV = GPHOV=='hov'

            # skip hov if no hov commute in Commute object
            if HOV and not self.hov:
                continue
            
            # Define key of column name, and output path for plot types
            key = {'hov_cong' : ['% cong (45.0)', '7_HOVCongestion'],
                  'gp_cong' : ['% cong (45.0)', '3_Congestion'],
                  'gp_sc' : ['% cong (36.0)', '4_SevereCongestion'],
                  'gp_tt' : ['Avg. TTS', '6_TravelTime'],
                  'gp_spd' : ['Avg. Speed', '8_Speed']}

            if ptype not in key.keys():
                raise Exception('ptype must be one of %s'%(key.keys()))

            #If necessary, create subfolder in plot directory
            line_path = os.path.join(self.paths['plot_path'], key[ptype][1])
            if not os.path.exists(line_path):
                os.mkdir(line_path)
            
            # Initiate plot
            w, h = [3.58, 1.1]
            lt, rt, bm, tp = [0.085, 0.97, 0.12, 0.8]
            fig, ax = plt.subplots(figsize=(w, h))
            fig.subplots_adjust(left=lt, right=rt, bottom=bm, top=tp)

            # Set plot parameters
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

            # set gridlines, ticks, and invisible borders
            ax.grid(axis = 'y', lw = 0.25)
            ax.tick_params(axis = 'x', direction = 'in')
            ax.yaxis.set_tick_params(color = '1', pad=0)
            ax.spines['left'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.spines['top'].set_visible(False)
            ax.set_xticks([5*12, 8*12, 11*12, 14*12, 17*12, 20*12])
            ax.set_xticklabels(['5 AM', '8 AM', '11 AM', '2 PM', '5 PM', '8 PM'])
            ax.set_xlim(left = 5*12, right = 20*12)

            # If applicable, set percentage y axis
            if measure in ['sc', 'cong']:
                ax.set_yticks([1,0.8,0.6,0.4,0.2])
                ax.set_yticklabels(['100%', '80%', '60%', '40%', '20%'])
                ax.set_ylim(bottom = 0, top = 1.01)        

            # Plot measures          
            if HOV:
                # If applicable, plot REV in dashed gray
                if self.rev:
                    ax.plot(self.current_rev_tt[key[ptype][0]],
                           color = '0.6',
                           lw = 0.75,
                           ls = '--')                

                # If applicable, plot GP(HOV), else plot GP in solid gray
                if self.gphov:
                    ax.plot(self.current_gphov_tt[key[ptype][0]],
                           color = '0.5',
                           lw = 0.75)
                else:
                    ax.plot(self.current_gp_tt[key[ptype][0]],
                           color = '0.5',
                           lw = 0.75)

                # Plot HOV in solid black
                ax.plot(self.current_hov_tt[key[ptype][0]],
                        color = '0',
                        lw = 0.75)
            else:
                # GP
                ax.plot(self.base_gp_tt[key[ptype][0]],
                       color = '0.5',
                       lw = 0.75)
                ax.plot(self.current_gp_tt[key[ptype][0]],
                       color = '0',
                       lw = 0.75)

            # Add title to plot
            name = self.name
            title = '%s %s%s'%(name[0:48] + '..' if len(name) > 50 else name,
                               'HOV ' if HOV else '',
                               measure)

            ax.set_title(title, fontsize = 8, loc = 'left')

            # save figure
            fname = '%s_%s%s.pdf'%(name.replace('/',' '), 
                                   'HOV' if HOV else '', 
                                   measure)
            f_path = os.path.join(line_path,fname)

            
            # Save figure
            try:
                pdffig = PdfPages(f_path)
            except IOError:
                now = datetime.datetime.now()
                now = now.strftime("%Y%m%d-%H%M%S")            
                f_path = f_path.replace('.pdf', ' %s.pdf'%now)
                pdffig = PdfPages(f_path)
            
            fig.savefig(pdffig, format = 'pdf')

            # write pdf properties
            properties = pdffig.infodict()
            properties['Title'] = \
                '%s %s%s%s %s'%(name, 
                                measure, 
                                ' HOV ' if HOV else ' ',
                                '' if HOV else self.base_year,
                                self.current_year)

            properties['Author'] = self.analyst + ' (Generated by PyMAS)'

            pdffig.close()
            plt.close()
        
        return None
            
    #---------------------------------------------------------------------------
    def comp_loops(self): 
        
        md = [x for x in self.__dict__.keys() if x.endswith('_md')]
        gphovs = map(lambda x: x.split('_')[1], md)
        
        for gphov in gphovs: 
            # Compare included loops
            b = getattr(self, 'base_%s_loops'%gphov)
            c = getattr(self, 'current_%s_loops'%gphov)

            b_l = b[b.map(lambda x: x not in c.values)]
            c_l = c[c.map(lambda x: x not in b.values)]

            b_l.reset_index(drop = True, inplace = True)
            c_l.reset_index(drop = True, inplace = True)

            b_l.name = 'in_base'
            c_l.name = 'in_current'

            setattr(self, '%s_loops_comp'%gphov, pd.concat([b_l, c_l], axis = 1))

        return None

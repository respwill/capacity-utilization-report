# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 11:18:43 2019

@author: parkhi
"""

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import datetime as dt
import time as tm
from tqdm import tqdm

class db_create():
    def __init__(self, file, factory, target_month, common_cc):
        
        dtnow = dt.datetime.fromtimestamp(tm.time())

        if dtnow.month != 1:
            self.year = dtnow.year
        else:
            self.year = dtnow.year - 1


        self.factory = factory
        
        self.common_cc = pd.read_excel(common_cc)
        self.fi_cc = list(self.common_cc['FI Cost Center'].drop_duplicates())
        
        self.raw_df = pd.read_excel(file, index_col='Category.1')
        
        self.raw_df = self.raw_df.rename_axis("Category")
        
        delete_month_list = [f'{self.year}.{str(x).zfill(2)}' for x in range(target_month+1, 13)]
        self.raw_df[delete_month_list] = 0
            
        self.cwd = os.getcwd()
        if os.path.isdir(self.cwd + "/result/{}".format(self.year)):
            pass
        else:
            os.makedirs(self.cwd + "/result/{}".format(self.year))
            
        self.K451_4185_counter = 0
        self.K451_4288_df = pd.DataFrame()
        self.K451_4185_df = pd.DataFrame()
        
# K4
        self.e_sip = []
        self.e_line = []
        self.f_line = []
        self.casip = []
        self.c_line = []
        self.scsp = []
        self.smlf = []
        self.b_line = []
        self.dps = []
        self.bumping = []
        self.g_line = []
# K3
        self.main_line = []
# K5
        self.fcbga = []
        self.cow = []
        self.sip = []
# Common
        self.etc = []

# K3,K4,K5 Labor
        self.k3_labor = []
        self.k4_labor = []
        self.k5_labor = []
        
# K3 production team 1 part includes 2 parts workers
        self.K350101100_counter = 0
        self.K350101100_df = pd.DataFrame()
        self.K370004500_df = pd.DataFrame()
        


# Line name collect    
#        self.line_name = ['E SIP', 'E Line', 'F Line', 'CASIP', 'C Line', 'SCSP', 'sMLF', 'B Line', 'DPS', 'Bumping', 'G Line', 'ETC']
        self.line_name = []

        
    def make_hour(self, data):
        return data / 3600
    
    def lab_drawing(self, last_month=12):
        print("="*50)
        print("LAB capacity utilization report".format(self.factory))
        print("="*50)
        
        
        month_list = ['{}.{}'.format(self.year, str(x).zfill(2)) for x in range(1, last_month+1)]
        month_list = ['Cost Center','Cost Center Description'] + month_list
        
        self.raw_df = self.raw_df[month_list]
        
        # cc_list = self.raw_df['Cost Center'].tolist()
        cc_list = self.raw_df['Cost Center'].drop_duplicates().tolist()

        for cc in tqdm(cc_list):
            
            cc_df = self.raw_df.loc[self.raw_df['Cost Center'] == cc]
            cc_desc = str(cc_df['Cost Center Description'].loc['Act Activity Qty'])
            
# only, product team related cc is need to be visualized.
            
            if (cc_desc.lower().find("Prod".lower()) != -1) and (cc_desc.lower().find("Tm".lower()) != -1) and (cc_desc.find(self.factory) == 0):
                pass
            
            # Bump CC is seprated and still need to be plotted.
            elif (cc_desc.lower().find("Prod".lower()) != -1) and (cc_desc.lower().find("Tm".lower()) != -1) and (cc_desc.find('Bump') == 0) and (cc.find(self.factory) == 0):
                pass
            
            # System module CC also need to be included.
            elif (cc_desc.lower().find("Prod".lower()) != -1) and (cc_desc.find('SM') == 0) and (cc[:2] == self.factory):
                pass
            
            # K5 CC description rule is different
            elif (cc_desc.lower().find("Production".lower()) != -1) and (cc_desc.lower().find("Pt".lower()) != -1) and (cc.find(self.factory) == 0):
                pass

            else:
                print("This cc doesn't need to be visualized")
                continue
        
            
            cc_df = cc_df[['{}.{}'.format(self.year, str(x).zfill(2)) for x in range(1,last_month+1)]]
            cc_df.columns= ['{}'.format(str(x).zfill(2)) for x in range(1,last_month+1)]
            cc_df = cc_df.T
            cc_df[['Act Activity Qty', 'Act Capa']] = cc_df[['Act Activity Qty', 'Act Capa']] / 3600
            cc_df = cc_df.round(2)

# K350101100 must combined with K370004500
            if cc == 'K350101100':
                self.K350101100_df = cc_df
            else:
                pass
                
            if cc == 'K370004500':
                self.K370004500_df = cc_df
            else:
                pass
                    
                
# if 'K350101100' comes first then save 'K350101100' and wait for 'K370004500'
            if (cc == 'K350101100') and (self.K370004500_df.empty):
                continue

# if 'K370004500' comes first and then 'K350101100' comes next            
            elif (self.K350101100_df.empty == False) and (self.K370004500_df.empty == False) and (self.K350101100_counter == 0):     
                self.K350101100_df[['Act Capa', 'Act Headcount']] = self.K350101100_df[['Act Capa', 'Act Headcount']] + self.K370004500_df[['Act Capa', 'Act Headcount']]
                self.K350101100_df[['Act Capa', 'Act Headcount']] = self.K350101100_df[['Act Capa', 'Act Headcount']].round(2)
                self.K350101100_df['Act Capa Util (%)'] = round((self.K350101100_df['Act Activity Qty'] / self.K350101100_df['Act Capa']) * 100)
                self.K350101100_df['Act Capa Util (%)'].fillna(0, inplace=True)
                
                fig = plt.figure()
                ax = fig.add_subplot(111)
                
                # need to change Title of fig as line name so that people can understand easily.
#                ax1 = self.K350101100_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title='{} ({}) Labor Capa Utilization in {}'.format('K350101100', 'K3 Test Prod Tm 1Pt [includes 2Pt]', self.year), kind='bar', legend=True )
                ax1 = self.K350101100_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title=f'K3 EL line / K350101100 (K3 Test Prod Tm 1Pt [includes 2Pt]) Labor Capa Utilization in {self.year}', kind='bar', legend=True )
                ax2 = self.K350101100_df[['Act Capa Util (%)']].plot(ax=ax1, secondary_y=True, color='purple', table=self.K350101100_df.T, kind='line', figsize=(14, 7), legend=True, linestyle=':')
                
                # x,y both must be int.
                x = [int(n)-1 for n in self.K350101100_df[['Act Capa Util (%)']].T.columns]
                y = self.K350101100_df[['Act Capa Util (%)']].T.values[0]

# annotate and if y value is bigger than 100, express as red color.
                for xy in zip(x,y):
                    if xy[1] > 100:
                        ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, color='red', fontsize='large', fontweight='bold')
                    else:
                        ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, fontsize='large', fontweight='bold')
                        
                        # Shrink current axis's height by 20% on the bottom
                box = ax.get_position()
                ax.set_position([box.x0, box.y0 + box.height * 0.2, box.width, box.height * 0.8])

# get labels and lines from each ax
                lines, labels = ax.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                
                lines = lines + lines2
                labels = labels + labels2

# Put a legend below current axis            
#            ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.2), fancybox=True, shadow=True, ncol=5)
                ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.2), ncol=5)
                
                
                self.k3_labor.append(fig)
                plt.close(fig)
                print("K3 labor fig saved")
                
                
                self.K350101100_counter += 1
                continue
                
            
            else:
                pass



            fig = plt.figure()
            ax = fig.add_subplot(111)

            # need to change Title of fig as line name so that people can understand easily.
            # make dictionary for cc and line name mapping
            cc_line = {'K350101100':'K3 EL line', 'K350101200':'K3 BE line', 'K350101300':'K3 E line', 'K460006400':'DPS, WBG', 
                       'K460004100':'SCSP, Saw MLF, CABGA FOL', 'K460001100':'SCSP, CABGA SMT', 'K460002100':'CABGA, WBSIP WB & ENCAP', 
                       'K460002200':'CABGA, WBSIP', 'K460004000':'FCSIP, FCSCSP, FCCSP', 'K460005100':'SCSP, Saw MLF, CABGA EOL', 
                       'K460006000':'FCCSP, FCSCSP', 'K460006100':'FCBGA, FPFCBGA, FCMBGA, COC FOL', 'K460006200':'FCBGA, FPFCBGA, FCMBGA, COC EOL',
                       'K460006300':'Wafer Bumping, WLCSP', 'K460006500':'FCTMVE', 'K470001100':'Mep, TCNCP', 'K560001110':'WLFO, SWIFT, COW, FCBGA',
                       'K560001200':'DPS', 'K560001400':'', 'K560001500':'SIP', 'K560001700':'', 'K560001300':'K5 support'}            

            ax1 = cc_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title=f'{cc_line[cc]} / {cc} ({cc_desc}) Labor Capa Utilization in {self.year}', kind='bar', legend=True )
            ax2 = cc_df[['Act Capa Util (%)']].plot(ax=ax1, secondary_y=True, color='purple', table=cc_df.T, kind='line', figsize=(14, 7), legend=True, linestyle=':')

# x,y both must be int.
            x = [int(n)-1 for n in cc_df[['Act Capa Util (%)']].T.columns]
            y = cc_df[['Act Capa Util (%)']].T.values[0]

# annotate and if y value is bigger than 100, express as red color.
# annotate on line plot
            for xy in zip(x,y):
                if xy[1] > 100:
                    ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, color='red', fontsize='large', fontweight='bold')
                else:
                    ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, fontsize='large', fontweight='bold')
            
# Shrink current axis's height by 20% on the bottom
            box = ax.get_position()
            ax.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 0.9])

# get labels and lines from each ax
            lines, labels = ax.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            
            lines = lines + lines2
            labels = labels + labels2

# Put a legend below current axis            
#            ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.2), fancybox=True, shadow=True, ncol=5)
            ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.175), ncol=5)
            
            
            if self.factory == 'K3':            
                self.k3_labor.append(fig)
                plt.close(fig)
                self.line_name.append('K3 Labor')
                print("K3 labor fig saved")
                    
            elif self.factory == 'K4':
                self.k4_labor.append(fig)
                plt.close(fig)
                self.line_name.append('K4 Labor')
                print("K4 labor fig saved")
            else:
                self.k5_labor.append(fig)
                plt.close(fig)
                self.line_name.append('K5 Labor')
                print("K5 labor fig saved")
                
        
        self.line_name = list(set(self.line_name))

# Send figures to PDF file by figure group.

        for line in tqdm(self.line_name):
            
            with PdfPages(self.cwd + '/result' +'/{}/{} Capacity Utilization in {}.pdf'.format(self.year, line, self.year)) as pdf:
                if line == 'K3 Labor':
                    pdf.attach_note("K3 Labor")
                    for idx, fig in enumerate(self.k3_labor): 
                        pdf.savefig(fig)
                        data_len = len(str(len(self.k3_labor)))
                        print("For {}, {} out of ".format(line, str(idx + 1).zfill(data_len)) + str(len(self.k3_labor)) + " exported to PDF")
                
                elif line == 'K4 Labor':
                    pdf.attach_note("K4 Labor")
                    for idx, fig in enumerate(self.k4_labor): 
                        pdf.savefig(fig)
                        data_len = len(str(len(self.k4_labor)))
                        print("For {}, {} out of ".format(line, str(idx + 1).zfill(data_len)) + str(len(self.k4_labor)) + " exported to PDF")
                else:
                    pdf.attach_note("K5 Labor")
                    for idx, fig in enumerate(self.k5_labor):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.k5_labor)))
                        print("For {}, {} out of ".format(line, str(idx + 1).zfill(data_len)) + str(len(self.k5_labor)) + " exported to PDF")
        
    
    def mach_drawing(self, last_month=12):
        print("="*50)
        print("{} Factory Mach capacity utilization report".format(self.factory))
        print("="*50)
        month_list = ['{}.{}'.format(self.year, str(x).zfill(2)) for x in range(1, last_month+1)]
        month_list = ['Cost Center','Cost Center Description'] + month_list
        
        self.raw_df = self.raw_df[month_list]
#        return self.raw_df
        
# separate data frame by cost center basis.
        factory_option = self.raw_df['Cost Center'].str.startswith(self.factory)
        self.raw_df = self.raw_df[factory_option]
        
        
        
        
        # some cc doesn't need to be included in this cc list.
        opt = self.raw_df['Cost Center Description'].str.contains('opt', case=False)
        incom = self.raw_df['Cost Center Description'].str.contains('incom', case=False)
        fvi = self.raw_df['Cost Center Description'].str.contains('fvi', case=False) | self.raw_df['Cost Center Description'].str.contains('f.v.i', case=False)
        packing = self.raw_df['Cost Center Description'].str.contains('packing', case=False)
        insp = self.raw_df['Cost Center Description'].str.contains('inspection', case=False) | self.raw_df['Cost Center Description'].str.contains('insp', case=False)
        drypack = self.raw_df['Cost Center Description'].str.contains('dry pack', case=False)
        dummy = self.raw_df['Cost Center Description'].str.contains('dummy', case=False)
        handler = self.raw_df['Cost Center Description'].str.contains('hd', case=False)
        #wfrmount = self.raw_df['Cost Center Description'].str.contains('wafer mount', case=False) | self.raw_df['Cost Center Description'].str.contains('wafr a mont', case=False)
        lami = self.raw_df['Cost Center Description'].str.contains('LAMINATION', case=False)
        wfrpack = self.raw_df['Cost Center Description'].str.contains('wafer pack', case=False)
        
        self.raw_df = self.raw_df.loc[~(opt| incom| fvi| packing| insp| drypack| dummy| handler| lami| wfrpack)]
        
        # other cc which hasn't been filtered by above.
        no_cc = ['K350101847', 'K350107094', 'K350108890', 'K420204005', 'K420204064', 'K411354005'
         'K411454005', 'K411454064', 'K420834037', 'K440114005', 'K440114064', 
         'K440115170', 'K440115729', 'K411355005', 'K411712265', 'K420704005', 
         'K420704064', 'K420804366', 'K420604005', 'K420604064', 'K430534036', 
         'K430534450', 'K511706412','K550401847', 'K520804035', 'K45010HD03', 
         
         # 2019.11.12 추가 
         'K430534013', 'K520801598', 'K520802200', 'K520802260',
         
         # 2019.12.06 추가 - SLT CC에 대해 Template allocation 문제가 있어 당분간 report에서 제외
         'K550407310',
         
         # 2020.01.10 추가 - K 관련 CC인데 현재로서는 유의미한 DATA가 아님
         'K540141992', 'K540144176'
         ]
        
        nontarget_cc = self.raw_df['Cost Center'].isin(no_cc)
        self.raw_df = self.raw_df.loc[~nontarget_cc]
        
        # excludes cc which tail is 4035 or 4130 or 4450 or 4490. they are no machine cc.
        
        no_cc_criteria = self.raw_df['Cost Center'].str.contains(r'(K4\d+4035|K4\d+4130|K4\d+4450|K4\d+4490)')
        self.raw_df = self.raw_df.loc[~no_cc_criteria]
        
        # excludes cc which are under common cc. in that case, checking FI cc is enough.
        
        lowlv_cc = list(self.common_cc['Cost Center'])
        low_cc = self.raw_df['Cost Center'].isin(lowlv_cc)
        self.raw_df = self.raw_df.loc[~low_cc]
        
        # make cc list with above filters.
        
        cc_list = self.raw_df['Cost Center'].drop_duplicates().tolist()

        for cc in tqdm(cc_list):
            
            cc_df = self.raw_df.loc[self.raw_df['Cost Center'] == cc]
            cc_desc = str(cc_df['Cost Center Description'].loc['Act Activity Qty'])
            

            cc_df = cc_df[['{}.{}'.format(self.year, str(x).zfill(2)) for x in range(1,last_month+1)]]
            cc_df.columns= ['{}'.format(str(x).zfill(2)) for x in range(1,last_month+1)]
            
            # need remove records which has no activity quantity and no capacity in target month.
            if (cc_df[f'{str(last_month).zfill(2)}'].loc['Act Activity Qty'] == 0) & (cc_df[f'{str(last_month).zfill(2)}'].loc['Act Capa'] == 0):
                print(f'Skip {cc} since it has no activity quantity and capacity' )
                continue
            else:
                pass
            
            
            cc_df = cc_df.T
            cc_df[['Act Activity Qty', 'Act Capa']] = cc_df[['Act Activity Qty', 'Act Capa']] / 3600
            cc_df = cc_df.round(2)
            
# K451_4185 must combined with K451_4288
            if cc == 'K451_4288':
                self.K451_4288_df = cc_df
            else:
                pass
                
            if cc == 'K451_4185':
                self.K451_4185_df = cc_df
            else:
                pass
                    
                
# if 'K451_4185' comes first then save 'K451_4185' and wait for 'K451_4288'
            if (cc == 'K451_4185') and (self.K451_4288_df.empty):
                continue

# if 'K451_4288' comes first and then 'K451_4185' comes next            
            elif (self.K451_4185_df.empty == False) and (self.K451_4288_df.empty == False) and (self.K451_4185_counter == 0):     
                self.K451_4185_df[['Act Activity Qty', 'Act Capa', 'Act Mach Qty']] = self.K451_4185_df[['Act Activity Qty', 'Act Capa', 'Act Mach Qty']] + self.K451_4288_df[['Act Activity Qty', 'Act Capa', 'Act Mach Qty']]
                self.K451_4185_df['Act Capa Util (%)'] = (self.K451_4185_df['Act Activity Qty'] / self.K451_4185_df['Act Capa']) * 100
                self.K451_4185_df['Act Capa Util (%)'].fillna(0, inplace=True)
                self.K451_4185_df = self.K451_4185_df.round(2)
                
                fig = plt.figure()
                ax = fig.add_subplot(111)
                
                # needs to mention about combine and low level cc.
                common_key = self.common_cc['FI Cost Center'] == 'K451_4185'
                common_key2 = self.common_cc['FI Cost Center'] == 'K451_4288'
                
                low_cc = list(self.common_cc['Cost Center'].loc[common_key])
                low_cc2 = list(self.common_cc['Cost Center'].loc[common_key2])
                final_low_cc = low_cc + low_cc2
                
                ax1 = self.K451_4185_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title='{} ({}) Machine Capa Utilization in {} \nCombined with K451_4288 *Common_cc: {}'.format('K451_4185', 'COAT & DEVELOP', self.year, final_low_cc), kind='bar', legend=True )
                ax2 = self.K451_4185_df[['Act Capa Util (%)']].plot(ax=ax1, secondary_y=True, color='purple', table=self.K451_4185_df.T, kind='line', figsize=(14, 7), legend=True, linestyle=':')
                
                # x,y both must be int.
                x = [int(n)-1 for n in self.K451_4185_df[['Act Capa Util (%)']].T.columns]
                y = self.K451_4185_df[['Act Capa Util (%)']].T.values[0]

# annotate and if y value is bigger than 100, express as red color.
                for xy in zip(x,y):
                    if xy[1] > 100:
                        ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, color='red', fontsize='large', fontweight='bold')
                    else:
                        ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, fontsize='large', fontweight='bold')
                        
# Shrink current axis's height by 20% on the bottom
                box = ax.get_position()
                ax.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 0.9])

# get labels and lines from each ax
                lines, labels = ax.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
            
                lines = lines + lines2
                labels = labels + labels2

# Put a legend below current axis            
#            ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.2), fancybox=True, shadow=True, ncol=5)
                ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.175), ncol=5)
                
                self.bumping.append(fig)
                plt.close(fig)
                print("E SIP fig saved")
                
                
                self.K451_4185_counter += 1
                continue
                
            
            else:
                pass
            
            
            fig = plt.figure()
            ax = fig.add_subplot(111)
            
            #it is better to mention low level cc if it is common cc
            if cc in self.fi_cc:
                common_key = self.common_cc['FI Cost Center'] == cc
                low_cc = list(self.common_cc['Cost Center'].loc[common_key])
                if len(low_cc) >= 10:
                    ax1 = cc_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title='{} ({}) Machine Capa Utilization in {} \n *Common_CC: {} \n{}'.format(cc, cc_desc, self.year, low_cc[:5], low_cc[5:]), kind='bar', legend=True )
                    ax2 = cc_df[['Act Capa Util (%)']].plot(ax=ax1, secondary_y=True, color='purple', table=cc_df.T, kind='line', figsize=(14, 7), legend=True, linestyle=':')
                else:
                    ax1 = cc_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title='{} ({}) Machine Capa Utilization in {} \n *Common_CC: {}'.format(cc, cc_desc, self.year, low_cc), kind='bar', legend=True )
                    ax2 = cc_df[['Act Capa Util (%)']].plot(ax=ax1, secondary_y=True, color='purple', table=cc_df.T, kind='line', figsize=(14, 7), legend=True, linestyle=':')
            else:
                ax1 = cc_df[['Act Activity Qty', 'Act Capa']].plot(ax=ax, title='{} ({}) Machine Capa Utilization in {}'.format(cc, cc_desc, self.year), kind='bar', legend=True )
                ax2 = cc_df[['Act Capa Util (%)']].plot(ax=ax1, secondary_y=True, color='purple', table=cc_df.T, kind='line', figsize=(14, 7), legend=True, linestyle=':')
            
            
# x,y both must be int.
            x = [int(n)-1 for n in cc_df[['Act Capa Util (%)']].T.columns]
            y = cc_df[['Act Capa Util (%)']].T.values[0]

# annotate and if y value is bigger than 100, express as red color.
# annotate on line plot
            for xy in zip(x,y):
                if xy[1] > 100:
                    ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, color='red', fontsize='large', fontweight='bold')
                else:
                    ax2.annotate('{} %'.format(xy[1]), xy=xy, xycoords='data', annotation_clip = True, fontsize='large', fontweight='bold')
            
# Shrink current axis's height by 20% on the bottom
            box = ax.get_position()
            ax.set_position([box.x0, box.y0 + box.height * 0.1, box.width, box.height * 0.9])

# get labels and lines from each ax
            lines, labels = ax.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            
            lines = lines + lines2
            labels = labels + labels2

# Put a legend below current axis            
#            ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.2), fancybox=True, shadow=True, ncol=5)
            ax.legend(lines, labels, loc='upper center', bbox_to_anchor=(0.5, -0.175), ncol=5)
            
            

# save each figure to each list
            

# Category is different by factory
                    
            if self.factory == 'K4':            
    # K4 E SIP
                
                if cc_desc.find('E SIP') == 0:
                    self.e_sip.append(fig)
                    plt.close(fig)
                    self.line_name.append('E SIP')
                    print("E SIP fig saved")
                    
    # K4 E LINE
                    
                elif ((cc_desc.find('E') == 0) and (cc_desc.find('E SIP') != 0)) or (cc_desc.find('K4_SCAN_TNR') ==0):
                    self.e_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('E Line')
                    print("E LINE fig saved")
                    
    # K4 F LINE
                    
                elif cc_desc.find('FCBGA') == 0:
                    self.f_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('F Line')
                    print("F LINE fig saved")
                    
    # K4 CASIP
                    
                elif cc_desc.find('CASIP') == 0:
                    self.casip.append(fig)
                    plt.close(fig)
                    self.line_name.append('CASIP')
                    print("CASIP fig saved")
                    
    # K4 C LINE
                    
                elif (cc_desc.find('FC CSP') == 0) or (cc_desc.find('FCCSP') == 0) or (cc_desc.find('MEP') == 0) or (cc_desc.find('BT FC POP') == 0) or (cc_desc.find('FC SCSP') == 0)  or\
                (cc_desc.find('FCHIP') == 0) or (cc_desc.find('FC POP') == 0) or (cc_desc.find('SIP') == 0):
                    self.c_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('C Line')
                    print("C LINE fig saved")
                    
                elif cc_desc in ['SAW/CLEAN2', 'UNDERFILL', 'S/BALL ATTACH2', 'PKG SAW3', 'O/S TEST2', 
                                 'VIA CLEAN', 'MOLD2','SIP S.MOUNT DEVIC FC', 'PLAS CLEAN12', 'UNDERFILL CU', 
                                 'PLASMA CLEAN3_2', 'UV IRRADIATE', 'K4 TCNCP', 'FLUX CLEAN', 'WAFER MOUNT2' ]:
                    self.c_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('C Line')
                    print("C LINE fig saved")
                    
    # K4 SCSP
                    
                elif (cc_desc.find('SCSP') == 0):
                    self.scsp.append(fig)
                    plt.close(fig)
                    self.line_name.append('SCSP')
                    print("SCSP fig saved")
                    
                elif (cc_desc.find('SMLF') != -1) or (cc_desc.find('FMLF') != -1):
                    self.smlf.append(fig)
                    plt.close(fig)
                    self.line_name.append('sMLF')
                    print("SMLF fig saved")
                    
    # K4 B LINE
                    
                elif (cc_desc.find('PBGA') == 0) or (cc_desc.find('PKG SAW') == 0) or (cc_desc.find('SBGA') == 0) or (cc_desc.find('POP') == 0): 
                    self.b_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('B Line')
                    print("B LINE fig saved")
                    
                elif cc_desc in ['DIE ATTACH', 'WIRE BOND', 'S.MOUNT DEVIC', 'MOLD', 'LASER MARK', 'S/BALL ATTACH', 'SAW/CLEAN', 
                                 'PKG SAW', 'DIMEN.INSP2', 'AOI insepection', 'O/S TEST', 'DIEATTACH CURE', 'PLASMA CLEAN', 'POST CURE',
                                 'BAKE FORDRYPACK', 'PLASMA CLEAN3', 'SINGULATION', 'WAFER MOUNT', 'BAKE FORDRYPACK', 'INK MARK',]:
                    self.b_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('B Line')
                    print("B LINE fig saved")
                    
    # K4 DPS
                    
                elif cc_desc.find('DPS') == 0:
                    self.dps.append(fig)
                    plt.close(fig)
                    self.line_name.append('DPS')
                    print("DPS fig saved")
                    
    # K4 BUMPING
                    
                elif cc_desc in ['SPUTTER 12"', 'PLATING 12"', 'TEMPLATE STRIP 12"', 'UBM SPUTTERING', 'ETCH', 'PLATING ETCH 12"', 
                                 'NI/SOLDER PLATG', 'ALIGN', 'HARD CURE', 'COAT & DEVELOP', 'DESCUM 12"', 'FLUX CLEANING 12"', 
                                 'PLASMA DESCUM', 'BUMP FLUX CLEANING', 'OVEN 12"', 'REFLOW', 'S/B ATT. ON WFR', 'BUMP CAP STRIP',
                                 'BUMP ACETIC CLEAN', 'Shear Test', 'Shear Test 12"', 'SORT WAFER 12"', 'SRD', 'SRD 12"', 'ACETIC ACID CLN 12"', 
                                 'Aligner 12"', 'FLUX CLEAN2', 'FLUX SCREEN PRT',  ]:
                    self.bumping.append(fig)
                    plt.close(fig)
                    self.line_name.append('Bumping')
                    print("BUMPING fig saved")
                    
    # K4 G LINE
                    
                elif cc_desc in ['WBGD', 'WFR EXPANSION', '1399_LASER CUT SAW']:
                    self.g_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('G Line')
                    print("G LINE fig saved")
                    
    # ETC
                    
                else:
                    self.etc.append(fig)
                    plt.close(fig)
                    self.line_name.append('ETC')
                    print("ETC fig saved")


            elif self.factory == 'K3':

# K3 Main CC                
                if cc_desc in ['K3_CATALYS', 'K3_CATALYS_RF', 'K3_HP93K_PS800', 'K3_HP93K_PS800_RF', 'K3_J750', 'K3_QUARTET', 'K3_HP93K_PS1600', 
                               'K3_HP93K_PS1600-RF', 'K3_110_PCT3', 'K3_110_ATS06', 'PKG STACK & REFLOW', 'K3_453_SLT', 'K3_SCAN', 'K3_TNR', 
                               'K3_UFLEX_COMMON', 'K3_T2000', 'K3_HP93K_C400', 'K3_HT3016',]:
                    self.main_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('MAIN')
                    print("Main line saved")
            
# K3 E CC
                elif (cc_desc.find('E') == 0) and (cc_desc.find('E SIP') != 0):
                    self.e_line.append(fig)
                    plt.close(fig)
                    self.line_name.append('E Line')
                    print("E LINE fig saved")
                    
                else:
                    self.etc.append(fig)
                    plt.close(fig)
                    self.line_name.append('ETC')
                    print("ETC fig saved")
                
            else:
# K5 DPS
                if (cc_desc.find('K5 DPS') == 0):
                    self.dps.append(fig)
                    plt.close(fig)
                    self.line_name.append('DPS')
                    print("K5 DPS fig saved")

# K5 COW               
                elif (cc_desc.find('K5 COW') == 0):
                    self.cow.append(fig)
                    plt.close(fig)
                    self.line_name.append('COW')
                    print("K5 COW fig saved")
                    
                elif cc_desc in ['K5 WAFER BACK GRIND', 'K5 SAW/CLEAN', 'K5 UV IRRADIATION', 'K5 LASER GROOVING', ]:
                    self.bumping.append(fig)
                    plt.close(fig)
                    self.line_name.append('Bumping')
                    print("K5 BUMP fig saved")
        
# K5 BUMP        
                elif (cc_desc.find('K5 BUMP') == 0):
                    self.bumping.append(fig)
                    plt.close(fig)
                    self.line_name.append('Bumping')
                    print("K5 BUMP fig saved")
                    
                elif cc_desc in ['K5 EDGE TRIM', 'K5 WAFER BOND', 'K5 DRY ETCH', 'K5 WAFER DEBOND', 'K5 SPUTTER', 
                                 'K5 COAT&DEVELOP', 'K5 MASK ALIGNER', 'K5 CURE PBO', 'K5 TEMPLATE STRIP', 'K5 SRD', 
                                 'K5 AOI', 'K5 DESCUM', 'K5 ETCH_SSAT', 'K5 ETCH_BATH', 'K5 PLATING', 'K5 REFLOW', 
                                 'K5 WAFER SORT', 'K5 WAFER FLIPPING', ]:
                    self.bumping.append(fig)
                    plt.close(fig)
                    self.line_name.append('Bumping')
                    print("K5 BUMP fig saved")
         
# K5 FCBGA
                elif (cc_desc.find('K5 FCBGA') == 0):
                    self.fcbga.append(fig)
                    plt.close(fig)
                    self.line_name.append('FCBGA')
                    print("K5 FCBGA fig saved")
                    
# K5 SIP                    
                elif (cc_desc.find('K5 SIP') == 0):
                    self.sip.append(fig)
                    plt.close(fig)
                    self.line_name.append('SIP')
                    print("K5 SIP fig saved")
                     
          
# K5 ETC
                else:
                    self.etc.append(fig)
                    plt.close(fig)
                    self.line_name.append('ETC')
                    print("ETC fig saved")
                    

# remove duplicated line name list
        self.line_name = list(set(self.line_name))

# Send figures to PDF file by figure group.

        for line in tqdm(self.line_name):
            
            with PdfPages(self.cwd + '/result' +'/{}/{} {} Machine Capacity Utilization in {}.pdf'.format(self.year, self.factory, line, self.year)) as pdf:
    
# K4 E SIP
                if line == 'E SIP':
                    pdf.attach_note("E SIP")
                    for idx, fig in enumerate(self.e_sip): 
                        pdf.savefig(fig)
                        data_len = len(str(len(self.e_sip)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.e_sip)) + " exported to PDF")
          
# K4 E LINE
                elif line == 'E Line':
                    pdf.attach_note("E Line")
                    for idx, fig in enumerate(self.e_line): 
                        pdf.savefig(fig)
                        data_len = len(str(len(self.e_line)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.e_line)) + " exported to PDF")
           
# K4 F LINE
                elif line == 'F Line':
                    pdf.attach_note("F Line")
                    for idx, fig in enumerate(self.f_line):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.f_line)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.f_line)) + " exported to PDF")
           
# K4 CASIP
                elif line == 'CASIP':
                    pdf.attach_note("CASIP")
                    for idx, fig in enumerate(self.casip):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.casip)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.casip)) + " exported to PDF")
                
# K4 C LINE
                elif line == 'C Line':
                    pdf.attach_note("C Line")
                    for idx, fig in enumerate(self.c_line):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.c_line)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.c_line)) + " exported to PDF")

# K4 SCSP
                elif line == 'SCSP':
                    pdf.attach_note("SCSP")
                    for idx, fig in enumerate(self.scsp):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.scsp)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.scsp)) + " exported to PDF")
         
# K4 SMLF
                elif line == 'sMLF':
                    pdf.attach_note("sMLF")
                    for idx, fig in enumerate(self.smlf):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.smlf)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.smlf)) + " exported to PDF")
         
# K4 B LINE
                elif line == 'B Line':
                    pdf.attach_note("B Line")
                    for idx, fig in enumerate(self.b_line):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.b_line)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.b_line)) + " exported to PDF")
           
# K4 DPS
                elif line == 'DPS':
                    pdf.attach_note("DPS")
                    for idx, fig in enumerate(self.dps):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.dps)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.dps)) + " exported to PDF")
        
# K4 BUMPING        
                elif line == 'Bumping':
                    pdf.attach_note("Bumping")
                    for idx, fig in enumerate(self.bumping):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.bumping)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.bumping)) + " exported to PDF")
                
# K4 G LINE
                elif line == 'G Line':
                    pdf.attach_note("G Line")
                    for idx, fig in enumerate(self.g_line):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.g_line)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.g_line)) + " exported to PDF")
 
# K3 MAIN LINE                       
                elif line == 'MAIN':
                    pdf.attach_note("Main Line")
                    for idx, fig in enumerate(self.main_line):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.main_line)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.main_line)) + " exported to PDF")
              
# K5 COW
                elif line == 'COW':
                    pdf.attach_note("COW Line")
                    for idx, fig in enumerate(self.cow):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.cow)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.cow)) + " exported to PDF")
        
# K5 FCBGA                
                elif line == 'FCBGA':
                    pdf.attach_note("FCBGA Line")
                    for idx, fig in enumerate(self.fcbga):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.fcbga)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.fcbga)) + " exported to PDF")
                    
                    
# K5 SIP                
                elif line == 'SIP':
                    pdf.attach_note("SIP Line")
                    for idx, fig in enumerate(self.sip):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.sip)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.sip)) + " exported to PDF")
           
# K3, K4, K5 ETC
                else:
                    pdf.attach_note("ETC")
                    for idx, fig in enumerate(self.etc):
                        pdf.savefig(fig)
                        data_len = len(str(len(self.etc)))
                        print("For {} {}, {} out of ".format(self.factory, line, str(idx + 1).zfill(data_len)) + str(len(self.etc)) + " exported to PDF")
    
    
# needs to analze capa to check if there is duplicated data.
class capaAnalyzer():
    def __init__(self, file, target_month):
        # coct center which has fi cc has no issue.
        self.capa = pd.read_excel('{}'.format(file))
        month_dict = {1:'January', 2:'February', 3:'March', 4:'April', 5:'May', 
                      6:'June', 7:'July', 8:'August', 9:'September', 
                      10:'October', 11:'November', 12:'December'}
        self.target_month = month_dict[target_month]
        
        dtnow = dt.datetime.fromtimestamp(tm.time())
        self.year = dtnow.year
        
    def check(self):
        cc_list = self.capa[['Cost Ctr','FI Cost Center']]
        
        no_fi_cc = pd.isna(self.capa['FI Cost Center'])
        no_fi_cc_capa = self.capa.loc[no_fi_cc]
        
        suspect_capa_raw = pd.merge(no_fi_cc_capa, cc_list, how='left', on='Cost Ctr')
        
        print(suspect_capa_raw.head(), suspect_capa_raw.shape)
        
        fi_cc = pd.notna(suspect_capa_raw['FI Cost Center_y'])
        suspect_capa = suspect_capa_raw.loc[fi_cc]
        cond1 = suspect_capa['Capacity of {}'.format(self.target_month)] != 0
        cond2 = suspect_capa['Total Quantity for {}'.format(self.target_month)] == 0
        suspect_capa_r = suspect_capa[cond1 & cond2]
        
        print(suspect_capa_r.head(), suspect_capa_r.shape)
        
        suspect_capa_r.to_excel('analyzing result_{}.{}.xlsx'.format(self.year, self.target_month))
        

mach_file = "Machine_act_raw.2019.01~12.XLSX"
lab_file = "Labor_act_raw_2019.01~12.XLSX"
target_month = 12
common_cc = 'common_cc.XLSX'
    
for site in ['K3','K4','K5']:
    mach = db_create(mach_file, site, target_month, common_cc)
    mach.mach_drawing()
    lab = db_create(lab_file, site, target_month, common_cc)
    lab.lab_drawing()
#    

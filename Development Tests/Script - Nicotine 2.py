#%% [markdown] 
# # Quantify Nicotine (wipe, dust, wristband)

#%% Setup and user variables
import pandas as pd
import numpy as np
from sklearn import linear_model
from sklearn.metrics import r2_score
import matplotlib.pyplot as plt
import altair as alt
from vega_datasets import data

# User variables
file_name = 'Data - Nicotine Template' # Input Excel filename
project_name = 'E.g., Cincinnati'
matrix_type = 'Wipe' # Wipe, Dust, Wristband, Pillow Case, Air Sample
analyte = 'Nicotine (ng/mL)' # Analyte concentration and units for export column title 
IS_Conc = 5 # Internal standard concentration (ng/mL)

#%% Initial full calibration: Quantification
# Data import
x = pd.read_excel(file_name + '.xlsx', sheet_name = 'Sheet1')

df_cal = x.loc[x['Type'] == 'Cal', ['SampleID', 'TC_Response', 'IS_Response', 'TC_Conc']]
df_cal['Response_Ratio'] = df_cal['TC_Response'] / df_cal['IS_Response']
df_cal['Conc_Ratio'] = df_cal['TC_Conc'] / IS_Conc

df_spl = x.loc[x['Type'] == 'Sample', ['SampleID', 'TC_Response', 'IS_Response']]
df_spl['Response_Ratio'] = df_spl['TC_Response'] / df_spl['IS_Response']

df_spl_MassHunter = x.loc[x['Type'] == 'Sample', ['SampleID', 'Analyte_CalcConc']]

# Linear model
x_train = df_cal.loc[:, 'Conc_Ratio'].values.reshape(-1,1) # format for lm
y_train = df_cal.loc[:, 'Response_Ratio'].values # format for lm
lm = linear_model.LinearRegression()
lm.fit(x_train, y_train)
 
# Calibration curve accuracy
df_cal['Measured_TC_Conc'] = ((df_cal['Response_Ratio'] - lm.intercept_) / lm.coef_) * IS_Conc
df_cal['Accuracy (%)'] = abs((((df_cal['Measured_TC_Conc'] - df_cal['TC_Conc']) / df_cal['TC_Conc']) * 100) - 100)
R_squared = np.round(r2_score(df_cal['Response_Ratio'], lm.predict(x_train)), decimals = 4)
Average_Cal_IS_Response = df_cal['IS_Response'].mean()

# Sample concentration
df_spl['Measured_TC_Conc'] = ((df_spl['Response_Ratio'] - lm.intercept_) / lm.coef_)  * IS_Conc
df_spl['Measured_Conc_Ratio'] = df_spl['Measured_TC_Conc'] / IS_Conc # for plot
df_spl['IS_Recovery'] = (df_spl['IS_Response'] / Average_Cal_IS_Response) * 100
Average_Spl_IS_Recovery = df_spl['IS_Recovery'].mean()
Average_Spl_TC_Conc = df_spl['Measured_TC_Conc'].mean()

# Check for and correct result for any TC peak area set to zero
df_spl.loc[df_spl['TC_Response'] == 0, ['Measured_TC_Conc']] = 0
df_spl.loc[df_spl['TC_Response'] == 0, ['Measured_Conc_Ratio']] = 0
# df_spl.head(3)

# Show table with calibration curve accuracy
def highlight_accuracy(s):
    '''
    Highlight poor accuracy values.
    '''
    is_accuracy = (s >= 110) | (s <= 90)
    return ['background-color: yellow' if v else '' for v in is_accuracy]

def highlight_difference(s):
    '''
    Highlight large percent differences.
    '''
    is_diff = (s >= 1)
    return ['background-color: orange' if v else '' for v in is_diff]

def highlight_set(s):
    '''
    Highlight calibration sets.
    '''
    is_set = (s == 'Set 1')
    return ['background-color: aqua' if v else 'rgb(255,228,181)' for v in is_set]

#%% Initial full calibration: Plot calibration curve
plt.rcParams.update(plt.rcParamsDefault) # reset theme
plt.style.use('ggplot') # applies to all plots in script
plt.figure(figsize=(5.5,5))
plt.scatter(df_spl['Measured_Conc_Ratio'], df_spl['Response_Ratio'], color = 'black', label = 'Sample') # Samples
plt.scatter(x_train, y_train, color = 'red', label = 'Standard') # Calibration points
plt.plot(x_train, lm.predict(x_train), color = 'red') # Calibration curve
plt.title('Initial Calibration Curve')
plt.suptitle(file_name, x = 0.05, y = 0.01, ha = 'left', size = 10)
plt.xlabel('TC/IS Concentration Ratio')
plt.ylabel('TC/IS Respose Ratio') 
plt.legend()
plt.annotate('$R^2$ = ' + str(R_squared), xy = (0.15, 0.75), xycoords = 'figure fraction')
plt.savefig('Graph1.png', dpi = 600, bbox_inches = 'tight')
plt.close()

import base64
data_uri = base64.b64encode(open('Graph1.png', 'rb').read()).decode('utf-8')
img_tag1 = '<img src="data:image/png;base64, {0}" style="width: 50%"/>'.format(data_uri)

#%% Initial full calibration: Plot IS recovery and TC concentration
fig_width = len(df_spl) * 0.3 # figure size based on number of samples
plt.figure(figsize=(fig_width, 7))

# IS recovery plot
plt.subplot(2, 1, 1)
plt.subplots_adjust(hspace = 0.5)
plt.scatter(df_spl['SampleID'], df_spl['IS_Recovery'], color = 'black', label = str())
plt.axhline(y = Average_Spl_IS_Recovery, color = 'blue', label = 'Sample mean')
plt.axhline(y = 50, color = 'orange', label = '50% recovery')
plt.title('Sample IS Recovery')
plt.xticks(rotation = 90)
plt.xlabel('Sample')
plt.ylabel('IS recovery (%)')
plt.legend(loc = 2)

# TC concentration plot
plt.subplot(2, 1, 2)
plt.scatter(df_spl['SampleID'], df_spl['Measured_TC_Conc'], color = 'dodgerblue', label = str())
plt.axhline(y = Average_Spl_TC_Conc, color = 'black', label = 'Sample mean TC concentration')
plt.title('Initial Sample TC Concentration')
plt.xticks(rotation = 90)
plt.xlabel('Sample')
plt.ylabel('TC concentration (ng/mL)')
plt.legend(loc = 2)
plt.suptitle(file_name, x = 0.1, y = 0, ha = 'left', size = 10)
plt.savefig('Graph2.png', dpi = 600, bbox_inches = 'tight')
plt.close()

data_uri = base64.b64encode(open('Graph2.png', 'rb').read()).decode('utf-8')
img_tag2 = '<img src="data:image/png;base64, {0}" style="width: 85%"/>'.format(data_uri)

#%% Split data to two sets
# Split calibration data based on accuracy of calibration points

# Alternative using position. Consider position if < 3 points in Set 1.
# df_cal_1 = df_cal.iloc[:6] 
# df_cal_1 = df_cal_1[['SampleID', 'TC_Response', 'IS_Response', 'TC_Conc', 'Response_Ratio', 'Conc_Ratio']]

# if len(df_cal_1) == 0:
#     raise ValueError('Corrected quantification unnecessary. Use alternate script.')

# df_cal_2 = df_cal.iloc[6:] 
# df_cal_2 = df_cal_2[['SampleID', 'TC_Response', 'IS_Response', 'TC_Conc', 'Response_Ratio', 'Conc_Ratio']]

# Using accuracy value
df_cal_1 = df_cal.loc[(df_cal['Accuracy (%)'] >= 110) | (df_cal['Accuracy (%)'] <= 90), 
    ['SampleID', 'TC_Response', 'IS_Response', 'TC_Conc', 'Response_Ratio', 'Conc_Ratio']]

if len(df_cal_1) == 0:
    raise ValueError('Corrected quantification unnecessary. Use alternate script.')

df_cal_2 = df_cal.loc[(df_cal['Accuracy (%)'] < 110) & (df_cal['Accuracy (%)'] > 90), 
    ['SampleID', 'TC_Response', 'IS_Response', 'TC_Conc', 'Response_Ratio', 'Conc_Ratio']]

Selection_cal = df_cal_1['Response_Ratio'].tail(1).to_numpy()
Selection_cal = Selection_cal.item(0) # use for selection in df_spl

# Split sample data based on TC_Response selection
df_spl_1 = df_spl.loc[df_spl['Response_Ratio'] <= Selection_cal, 
    ['SampleID', 'TC_Response', 'IS_Response', 'Response_Ratio']]
df_spl_2 = df_spl.loc[df_spl['Response_Ratio'] > Selection_cal, 
    ['SampleID', 'TC_Response', 'IS_Response', 'Response_Ratio']]

#%% Quantificaton of first set
# Linear model
x_train_1 = df_cal_1.loc[:, 'Conc_Ratio'].values.reshape(-1,1) # format for lm
y_train_1 = df_cal_1.loc[:, 'Response_Ratio'].values # format for lm
lm_1 = linear_model.LinearRegression()
lm_1.fit(x_train_1, y_train_1)

# Quantificaton of first set: Set 1 calibration curve accuracy
df_cal_1['Measured_TC_Conc'] = ((df_cal_1['Response_Ratio'] - lm_1.intercept_) / lm_1.coef_) * IS_Conc
df_cal_1['Accuracy (%)'] = abs((((df_cal_1['Measured_TC_Conc'] - df_cal_1['TC_Conc']) / df_cal_1['TC_Conc']) * 100) - 100)
R_squared_1 = np.around(r2_score(df_cal_1['Response_Ratio'], lm_1.predict(x_train_1)), decimals = 4)
Average_Cal_IS_Response_1 = df_cal_1['IS_Response'].mean()

# Quantificaton of first set: Set 1 sample concentration
df_spl_1['Measured_TC_Conc'] = ((df_spl_1['Response_Ratio'] - lm_1.intercept_) / lm_1.coef_)  * IS_Conc
df_spl_1['Measured_Conc_Ratio'] = df_spl_1['Measured_TC_Conc'] / IS_Conc # for plot
df_spl_1['IS_Recovery'] = (df_spl_1['IS_Response'] / Average_Cal_IS_Response_1) * 100
# df_spl_1.head(3)

#%% Quantification of second set
# Linear model
x_train_2= df_cal_2.loc[:, 'Conc_Ratio'].values.reshape(-1,1) # format for lm
y_train_2 = df_cal_2.loc[:, 'Response_Ratio'].values # format for lm
lm_2 = linear_model.LinearRegression()
lm_2.fit(x_train_2, y_train_2)

# Quantification of second set: Set 2 calibration curve accuracy
df_cal_2['Measured_TC_Conc'] = ((df_cal_2['Response_Ratio'] - lm_2.intercept_) / lm_2.coef_) * IS_Conc
df_cal_2['Accuracy (%)'] = abs((((df_cal_2['Measured_TC_Conc'] - df_cal_2['TC_Conc']) / df_cal_2['TC_Conc']) * 100) - 100)
R_squared_2 = np.around(r2_score(df_cal_2['Response_Ratio'], lm_2.predict(x_train_2)), decimals = 4)
Average_Cal_IS_Response_2 = df_cal_2['IS_Response'].mean()

# Quantification of second set: Set 2 sample concentration
df_spl_2['Measured_TC_Conc'] = ((df_spl_2['Response_Ratio'] - lm_2.intercept_) / lm_2.coef_) * IS_Conc
df_spl_2['Measured_Conc_Ratio'] = df_spl_2['Measured_TC_Conc'] / IS_Conc # for plot
df_spl_2['IS_Recovery'] = (df_spl_2['IS_Response'] / Average_Cal_IS_Response_2) * 100
# df_spl_2.head(3)

#%% Secondary quantification: Plot calibration curves
plt.figure(figsize=(7.5,3.5))

plt.subplot(1, 2, 1)
plt.scatter(df_spl_1['Measured_Conc_Ratio'], df_spl_1['Response_Ratio'], color = 'black', label = 'Sample') # Samples
plt.scatter(x_train_1, y_train_1, color = 'red', label = 'Standard') # Calibration points
plt.plot(x_train_1, lm_1.predict(x_train_1), color = 'red') # Calibration curve
plt.title('Set 1 Full Scale Cal. Curve')
plt.suptitle(file_name, x = 0.05, y = -0.05, ha = 'left', size = 10)
plt.xlabel('TC/IS Concentration Ratio')
plt.ylabel('TC/IS Respose Ratio') 
plt.legend()
plt.annotate('$R^2$ = ' + str(R_squared_1), xy = (0.125, 0.7), xycoords = 'figure fraction')

plt.subplot(1, 2, 2)
plt.scatter(df_spl_2['Measured_Conc_Ratio'], df_spl_2['Response_Ratio'], color = 'black', label = 'Sample') # Samples
plt.scatter(x_train_2, y_train_2, color = 'red', label = 'Standard') # Calibration points
plt.plot(x_train_2, lm_2.predict(x_train_2), color = 'red') # Calibration curve
plt.title('Set 2 Full Scale Cal. Curve')
plt.xlabel('TC/IS Concentration Ratio') 
plt.legend()
plt.annotate('$R^2$ = ' + str(R_squared_2), xy = (0.625, 0.7), xycoords = 'figure fraction')

plt.savefig('Graph3.png', dpi = 600, bbox_inches = 'tight')
plt.close()
# plt.show()

data_uri = base64.b64encode(open('Graph3.png', 'rb').read()).decode('utf-8')
img_tag3 = '<img src="data:image/png;base64, {0}" style="width: 70%"/>'.format(data_uri)

#%% Export data
df_spl_1['Set'] = 'Set 1'
df_spl_2['Set'] = 'Set 2'

df_cal_1['Set'] = 'Set 1'
df_cal_2['Set'] = 'Set 2'

table_position_cal = len(df_cal_1.index)
table_position_unc = len(df_cal.index)

# Combine corrected sample data sets
df_spl_combined = df_spl_1.append(df_spl_2)

# Check for and correct result for any TC peak area set to zero
df_spl_combined.loc[df_spl_combined['TC_Response'] == 0, ['Measured_TC_Conc']] = 0
df_spl_combined.loc[df_spl_combined['TC_Response'] == 0, ['Measured_Conc_Ratio']] = 0
 
# Dataframe for primary export
df_spl_combined_export = df_spl_combined.loc[:, ['SampleID', 'IS_Recovery', 'Measured_TC_Conc']].sort_index()
df_spl_combined_export['Batch'] = file_name
df_spl_combined_export = df_spl_combined_export.loc[:, ['Batch', 'SampleID', 'IS_Recovery', 'Measured_TC_Conc']]

# Compare corrected to initial quantification results and MassHunter results
df_spl_for_merge = df_spl.loc[:, ['SampleID', 'Measured_TC_Conc']]
df_comparison = pd.merge(df_spl_combined_export, df_spl_for_merge, how='outer', on='SampleID')
df_comparison['Percent_Difference'] = (abs(df_comparison['Measured_TC_Conc_x'] - df_comparison['Measured_TC_Conc_y']) / df_comparison['Measured_TC_Conc_y']) * 100
df_comparison = pd.merge(df_comparison, df_spl_MassHunter, how='outer', on='SampleID')
df_comparison_sorted = df_comparison.sort_values(by=['Percent_Difference'], ascending=False)
df_comparison_sorted = df_comparison_sorted.rename(index=str, 
    columns={'Measured_TC_Conc_x':'Measured_TC_Conc_Corrected', 'Measured_TC_Conc_y':'Measured_TC_Conc_Uncorrected', 
    'Analyte_CalcConc':'Measured_TC_Conc_MassHunter'})
df_comparison_sorted = df_comparison_sorted.loc[:, ['SampleID', 'IS_Recovery', 'Measured_TC_Conc_Uncorrected',
    'Measured_TC_Conc_MassHunter','Measured_TC_Conc_Corrected', 'Percent_Difference']]

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(file_name + ' Results' + '.xlsx', engine='xlsxwriter')
df_spl_combined_export = df_spl_combined_export.rename(index=str, columns={'Measured_TC_Conc':analyte})
df_spl_combined_export.to_excel(writer, sheet_name='Export')
writer.save()

html_df_cal = df_cal \
    .round({'TC_Response' : 2, 
            'IS_Response' : 2, 
            'TC_Conc' : 1, 
            'Response_Ratio' : 3,
            'Measured_TC_Conc' : 3, 
            'Accuracy (%)' : 0}) \
    .style \
        .set_table_attributes('border="1" class="dataframe mystyle"') \
        .apply(highlight_accuracy, subset=['Accuracy (%)']) \
        .format({'TC_Response' : '{:.2f}',
                'IS_Response' : '{:.2f}',
                'Conc_Ratio' : '{:.2f}', 
                'TC_Conc' : '{:.1f}',
                'Response_Ratio' : '{:.3f}',
                'Measured_TC_Conc' : '{:.3f}',
                'Accuracy (%)' : '{:.0f}'}) \
        .render()

html_df_cal_1 = df_cal_1 \
    .round({'TC_Response' : 2, 
            'IS_Response' : 2, 
            'TC_Conc' : 1, 
            'Response_Ratio' : 3,
            'Measured_TC_Conc' : 3, 
            'Accuracy (%)' : 0}) \
    .style \
        .set_table_attributes('border="1" class="dataframe mystyle"') \
        .apply(highlight_accuracy, subset=['Accuracy (%)']) \
        .format({'TC_Response' : '{:.2f}',
                'IS_Response' : '{:.2f}',
                'Conc_Ratio' : '{:.2f}', 
                'TC_Conc' : '{:.1f}',
                'Response_Ratio' : '{:.3f}',
                'Measured_TC_Conc' : '{:.3f}',
                'Accuracy (%)' : '{:.0f}'}) \
        .render()

html_df_cal_2 = df_cal_2 \
    .round({'TC_Response' : 2, 
            'IS_Response' : 2, 
            'TC_Conc' : 1, 
            'Response_Ratio' : 3,
            'Measured_TC_Conc' : 3, 
            'Accuracy (%)' : 0}) \
    .style \
        .set_table_attributes('border="1" class="dataframe mystyle"') \
        .apply(highlight_accuracy, subset=['Accuracy (%)']) \
        .format({'TC_Response' : '{:.2f}',
                'IS_Response' : '{:.2f}',
                'Conc_Ratio' : '{:.2f}', 
                'TC_Conc' : '{:.1f}',
                'Response_Ratio' : '{:.3f}',
                'Measured_TC_Conc' : '{:.3f}',
                'Accuracy (%)' : '{:.0f}'}) \
        .render()

html_df_comparison_sorted = df_comparison_sorted \
    .round({'IS_Recovery' : 0,
            'Measured_TC_Conc_Corrected' : 3,
            'Measured_TC_Conc_Uncorrected' : 3,
            'Percent_Difference' : 2,
            'Measured_TC_Conc_MassHunter' : 3}) \
    .style \
        .set_table_attributes('border="1" class="dataframe mystyle"') \
        .apply(highlight_difference, subset=['Percent_Difference']) \
        .format({'IS_Recovery' : '{:.0f}',
                'Measured_TC_Conc_Corrected' : '{:.3f}',
                'Measured_TC_Conc_Uncorrected' : '{:.3f}',
                'Percent_Difference' : '{:.2f}',
                'Measured_TC_Conc_MassHunter' : '{:.3f}'}) \
        .render()

html_df_spl_combined = df_spl_combined \
    .round({'TC_Response' : 3,	
        'IS_Response' : 3,
        'Response_Ratio' : 3,	
        'Measured_TC_Conc' : 3,	
        'Measured_Conc_Ratio' : 3,	
        'IS_Recovery' : 0}) \
    .style \
        .set_table_attributes('border="1" class="dataframe mystyle"') \
        .apply(highlight_set, subset=['Set']) \
        .format({'TC_Response' : '{:.3f}',	
                'IS_Response' : '{:.3f}',
                'Response_Ratio' : '{:.3f}',	
                'Measured_TC_Conc' : '{:.3f}',	
                'Measured_Conc_Ratio' : '{:.3f}',	
                'IS_Recovery' : '{:.0f}'}) \
        .render()

#%% Table output
#pd.set_option('colheader_justify', 'center')   # FOR TABLE <th>

html_string = '''
<html>
  <head><title>Quantification Quality Control Report</title></head>
  <link rel="stylesheet" type="text/css" href="df_style.css"/>
  <body>
    <h1 class="title">Quality Control Report</h1>
    <h1>Sample Sequence: {header_info}</h1>
        <p>Project: {project_info}</p>
        <p>Matrix: {matrix_info}</p>
        <p>Analyte: {analyte_info}
        <p>Internal Standard Concentration: {IS_info} ng/mL</p>
        <p>Nicotine quantification is reported as ng/mL in the autosampler vial. The concentration needs to be converted to nicotine mass (ng) prior to submission.</p>
        <p>TC = Target Compound, IS = Internal Standard</p>
    <h2>Initial uncorrected calibration</h2>
        <h3>Uncorrected calibration curve accuracy</h3>
            {table3}
        <div class="box">
            <div class="one_alt">
                <h3>Sample distribution within uncorrected calibration range</h3>
                    <div class="one_alt2">
                    {plot1}
                    </div>
            </div>
            <div class="two_alt">
                <h3>Internal standard (IS) recovery and target compound (TC) concentration</h3>
                    <p>IS recovery vs. TC concentration over the injection sequence.</p>
                        {plot2}
            </div>
        </div>
    <h2>Corrected calibration</h2>
        <p>Data is split into two sets based on calibration point accuracy.</p>
        <div class="box">
        <div class="one">
            <h3>Calibration curve accuracy: Set 1</h3>
                {table4}
            <h3>Calibration curve accuracy: Set 2</h3>
                {table5}
        </div>
        <div class="two">
            <h3>Sample distribution within corrected calibration ranges</h3>
                {plot3}
        </div>
        </div>
     <div class="box">
     <div class="one">   
    <h2>Comparison of corrected and un-corrected data</h2>
        <p>Sorted by percent difference. Includes comparison with Mass Hunter uncorrected data.</p> 
                {table6}
        </div>
        <div class="two">
    <h2>Samples in each of the corrected calibration curve ranges</h2>
        <p>Sorted by set.</p>
                {table7}
        </div>
        </div>
  </body>
</html>
'''
with open('Draft QC Report.html', 'w') as f:
    f.write(html_string.format(table3=html_df_cal,
        table4=html_df_cal_1,
        table5=html_df_cal_2,
        table6=html_df_comparison_sorted,
        table7=html_df_spl_combined,
        plot1=img_tag1,
        plot2=img_tag2,
        plot3=img_tag3,
        header_info=file_name,
        project_info=project_name,
        matrix_info=matrix_type,
        IS_info=IS_Conc,
        analyte_info=analyte))


#%%

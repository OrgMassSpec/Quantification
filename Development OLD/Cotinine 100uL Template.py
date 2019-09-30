#%% [markdown] 
# # Quantify Cotinine (100 uL urine sample)
#
# Accounts for 20 fold dilution. Sample size is 100 uL of urine (20x lower than the 2 mL typical size), and add 1900 uL water to a total volume of 2 mL. Next spike IS and work up as normal (100 uL of 0.1 ng/uL cotinine-d3, or 10 ng cotinine-d3). After quantifying, multiply by 20 to get original concentration.
#
# Throughout code, TC = Target Compound, IS = Internal Standard

#%% Setup and user variables
import pandas as pd
import numpy as np
from sklearn import linear_model
from sklearn.metrics import r2_score
import matplotlib.pyplot as plt
import altair as alt
from vega_datasets import data

# User variables
file_name = 'Cotinine 100uL Template Test Data' # Input Excel filename
analyte = 'Cotinine Conc. (ng/mL)' # Analyte concentration and units for export column title 
IS_Conc = 5 # Internal standard concentration (ng/mL)

#%% [markdown]
## Initial full calibration

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
    return ['background-color: brown' if v else '' for v in is_accuracy]

df_cal.style.apply(highlight_accuracy, subset=['Accuracy (%)'])

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
plt.savefig(file_name + ' 1 Initial Calibration Curve.png', dpi = 600, bbox_inches = 'tight')
plt.close()

# Interactive calibration curve
# Dataframe containing calibration curve 
cal_plot_df = pd.DataFrame({
    'x': x_train.flatten(), 
    'y': y_train, 
    'y_predicted': lm.predict(x_train)
})

chart_spl = alt.Chart(df_spl).mark_circle(size=100, color='black').encode(
    x='Measured_Conc_Ratio',
    y='Response_Ratio',
    tooltip=['SampleID']
).interactive()

chart_cal = alt.Chart(cal_plot_df).mark_circle(size=100, color='red').encode(
    x='x',
    y='y',
    #color='Origin',
    #tooltip=['SampleID']
).interactive()

chart_fit = alt.Chart(cal_plot_df).mark_line(color='red').encode(
    x='x:Q',
    y='y_predicted:Q'
)

chart = chart_spl + chart_cal + chart_fit
chart.save(file_name + ' 1 Initial Calibration Curve.html')

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
plt.savefig(file_name + ' 2 IS Recovery and Initial TC Concentration.png', dpi = 600, bbox_inches = 'tight')
plt.close()

#%% [markdown]
# ## Split data
#
# Data is split into two sets based on: calibration point accuracy. 
#
# ** Option: calibraiton point accuracy or position. (Consider position if < 3 points in Set 1.)

#%% Split data to two sets
# Split calibration data based on accuracy of calibration points
# Using position
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

df_cal_1

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

df_cal_2

#%% Secondary quantification: Plot calibration curves
plt.figure(figsize=(7.5,3.5))

plt.subplot(1, 2, 1)
plt.scatter(df_spl_1['Measured_Conc_Ratio'], df_spl_1['Response_Ratio'], 
    color = 'black', label = 'Sample') # Samples
plt.scatter(x_train_1, y_train_1, color = 'red', label = 'Standard') # Calibration points
plt.plot(x_train_1, lm_1.predict(x_train_1), color = 'red') # Calibration curve
plt.title('Set 1 Full Scale Cal. Curve')
plt.suptitle(file_name, x = 0.05, y = -0.05, ha = 'left', size = 10)
plt.xlabel('TC/IS Concentration Ratio')
plt.ylabel('TC/IS Respose Ratio') 
plt.legend()
plt.annotate('$R^2$ = ' + str(R_squared_1), xy = (0.125, 0.7), xycoords = 'figure fraction')

plt.subplot(1, 2, 2)
plt.scatter(df_spl_2['Measured_Conc_Ratio'], df_spl_2['Response_Ratio'], 
    color = 'black', label = 'Sample') # Samples
plt.scatter(x_train_2, y_train_2, color = 'red', label = 'Standard') # Calibration points
plt.plot(x_train_2, lm_2.predict(x_train_2), color = 'red') # Calibration curve
plt.title('Set 2 Full Scale Cal. Curve')
plt.xlabel('TC/IS Concentration Ratio') 
plt.legend()
plt.annotate('$R^2$ = ' + str(R_squared_2), xy = (0.625, 0.7), 
    xycoords = 'figure fraction')

plt.savefig(file_name + ' 4 Corrected Calibration Curves.png', 
    dpi = 600, bbox_inches = 'tight')
plt.close()
# plt.show()

#%% [markdown]
# ## Export to Excel
# Correct for sample dilution 

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
df_spl_combined.loc[df_spl_combined['TC_Response'] == 0, 
    ['Measured_TC_Conc']] = 0
df_spl_combined.loc[df_spl_combined['TC_Response'] == 0, 
    ['Measured_Conc_Ratio']] = 0
 
# Dataframe for primary export
# Measured_TC_Conc is the concentration in the vial (the diluted form, ng/mL)
df_spl_combined_export = df_spl_combined.loc[:, ['SampleID', 
    'IS_Recovery', 'Measured_TC_Conc']].sort_index()
df_spl_combined_export['Batch'] = file_name
df_spl_combined_export['Cotinine in Urine (ng cotinine / mL urine)'] = df_spl_combined_export['Measured_TC_Conc'] * 20
df_spl_combined_export = df_spl_combined_export.loc[:, ['Batch', 
    'SampleID', 'IS_Recovery', 'Measured_TC_Conc', 
    'Cotinine in Urine (ng cotinine / mL urine)']]

# Compare corrected to initial quantification results and MassHunter results
df_spl_for_merge = df_spl.loc[:, ['SampleID', 'Measured_TC_Conc']]
df_comparison = pd.merge(df_spl_combined_export, df_spl_for_merge, 
    how='outer', on='SampleID')
df_comparison['Percent_Difference'] = ((abs(df_comparison['Measured_TC_Conc_x'] -
    df_comparison['Measured_TC_Conc_y']) / 
    df_comparison['Measured_TC_Conc_y']) * 100)
df_comparison = pd.merge(df_comparison, df_spl_MassHunter, 
    how='outer', on='SampleID')
df_comparison_sorted = df_comparison.sort_values(by=['Percent_Difference'], 
    ascending=False)
df_comparison_sorted = df_comparison_sorted.rename(index=str, 
    columns={'Measured_TC_Conc_x':'Measured_TC_Conc_Corrected', 
    'Measured_TC_Conc_y':'Measured_TC_Conc_Uncorrected', 
    'Analyte_CalcConc':'Measured_TC_Conc_MassHunter'})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(file_name + ' Results' + '.xlsx', engine='xlsxwriter')
df_spl_combined_export = df_spl_combined_export.rename(index=str, 
    columns={'Measured_TC_Conc':analyte})
df_spl_combined_export.to_excel(writer, sheet_name='Export')
df_comparison_sorted.to_excel(writer, sheet_name='Comparison')
df_spl_combined.to_excel(writer, sheet_name='Corrected_Samples')
df_cal_1.to_excel(writer, sheet_name='Corrected_Calibration')
df_cal_2.to_excel(writer, sheet_name='Corrected_Calibration', 
    startrow=table_position_cal + 2)
df_cal.to_excel(writer, sheet_name='Uncorrected_Data')
df_spl.to_excel(writer, sheet_name='Uncorrected_Data', 
    startrow=table_position_unc + 2)
writer.save()

#%% [markdown] 
# # Quantify creatinine
# TC = Target Compound, IS = Internal Standard

# Accounts for 10,000 fold dilution and reports urinary creatinine concentration (ug creatinine / mL urine).
 
#%% ########## Python imports, user variables 
import pandas as pd
import numpy as np
from sklearn import linear_model
from sklearn.metrics import r2_score
import matplotlib.pyplot as plt

# User variables
file_name = 'Data - Creatinine Template' # Input Excel filename
#analyte = 'Creatinine Conc. (ug/mL)' # Analyte concentration and units for export column title 
IS_Conc = 0.1 # Internal standard concentration (ug/mL)

#%% ########## Calibration: Data import (low end calibration correction unnecessary)
# Formatted in Excel prior to input.
x = pd.read_excel(file_name + '.xlsx', sheet_name = 'Sheet1')

df_cal = x.loc[x['Type'] == 'Cal', ['SampleID', 'TC_Response', 'IS_Response', 'TC_Conc']]
df_cal['Response_Ratio'] = df_cal['TC_Response'] / df_cal['IS_Response']
df_cal['Conc_Ratio'] = df_cal['TC_Conc'] / IS_Conc

df_spl = x.loc[x['Type'] == 'Sample', ['SampleID', 'TC_Response', 'IS_Response']]
df_spl['Response_Ratio'] = df_spl['TC_Response'] / df_spl['IS_Response']

df_spl_MassHunter = x.loc[x['Type'] == 'Sample', ['SampleID', 'Analyte_CalcConc']]

#%% ########## Calibration: Linear model
# Linear model
x_train = df_cal.loc[:, 'Conc_Ratio'].values.reshape(-1,1) # format for lm
y_train = df_cal.loc[:, 'Response_Ratio'].values # format for lm
lm = linear_model.LinearRegression()
lm.fit(x_train, y_train)

#%% ########## Calibration curve accuracy
df_cal['Measured_TC_Conc'] = ((df_cal['Response_Ratio'] - lm.intercept_) / lm.coef_) * IS_Conc
df_cal['Accuracy (%)'] = abs((((df_cal['Measured_TC_Conc'] - df_cal['TC_Conc']) / df_cal['TC_Conc']) * 100) - 100)
R_squared = np.around(r2_score(df_cal['Response_Ratio'], lm.predict(x_train)), decimals = 4)
Average_Cal_IS_Response = df_cal['IS_Response'].mean()
# df_cal.head(3)

#%% ########## Sample (autosampler vial) concentration
df_spl['Measured_TC_Conc'] = ((df_spl['Response_Ratio'] - lm.intercept_) / lm.coef_)  * IS_Conc
df_spl['Measured_Conc_Ratio'] = df_spl['Measured_TC_Conc'] / IS_Conc # for plot
df_spl['IS_Recovery'] = (df_spl['IS_Response'] / Average_Cal_IS_Response) * 100
Average_Spl_IS_Recovery = df_spl['IS_Recovery'].mean()
Average_Spl_TC_Conc = df_spl['Measured_TC_Conc'].mean()

# Check for and correct result for any TC peak area set to zero
df_spl.loc[df_spl['TC_Response'] == 0, ['Measured_TC_Conc']] = 0
df_spl.loc[df_spl['TC_Response'] == 0, ['Measured_Conc_Ratio']] = 0
# df_spl.head(3)

#%% ########## Plots
# Calibration curve (regression)
plt.rcParams.update(plt.rcParamsDefault) # reset theme
plt.style.use('ggplot') # applies to all plots in script

plt.figure(figsize=(5, 4.5))

plt.scatter(df_spl['Measured_Conc_Ratio'], df_spl['Response_Ratio'], color = 'black', label = 'Sample') # Samples
plt.scatter(x_train, y_train, color = 'red', label = 'Standard') # Calibration points
plt.plot(x_train, lm.predict(x_train), color = 'red') # Calibration curve
plt.title('Creatinine Calibration Curve (Vial)')
plt.suptitle(file_name, x = 0.05, y = 0.01, ha = 'left', size = 10)
plt.xlabel('TC/IS Concentration Ratio')
plt.ylabel('TC/IS Respose Ratio') 
plt.legend()
plt.annotate('$R^2$ = ' + str(R_squared), xy = (0.15, 0.75), xycoords = 'figure fraction')

plt.savefig(file_name + ' 1 Calibration Curve.png', dpi = 600, bbox_inches = 'tight')
plt.close()
# plt.show()

#%% IS recovery plot
plt.figure(figsize=(10,3.5))
plt.scatter(df_spl['SampleID'], df_spl['IS_Recovery'], color = 'black', label = str())
plt.axhline(y = Average_Spl_IS_Recovery, color = 'blue', label = 'Sample mean')
plt.axhline(y = 50, color = 'orange', label = '50% recovery')
plt.title('Sample IS Recovery')
plt.xticks(rotation = 90)
plt.xlabel('Sample')
plt.ylabel('IS recovery (%)')
plt.legend(loc = 2)
plt.suptitle(file_name, x = 0.1, y = -0.175, ha = 'left', size = 10)

plt.savefig(file_name + ' 2 IS Recovery.png', dpi = 600, bbox_inches = 'tight')
plt.close()
# plt.show()

#%% TC concentration plot
plt.figure(figsize=(10,3.5))
plt.scatter(df_spl['SampleID'], df_spl['Measured_TC_Conc'], color = 'dodgerblue', label = str())
plt.axhline(y = Average_Spl_TC_Conc, color = 'black', label = 'Sample mean TC concentration')
plt.title('Batch Creatinine Concentration (ug creatinine / mL diluted urine)')
plt.xticks(rotation = 90)
plt.xlabel('Sample')
plt.ylabel('Concentration (ug / diuluted urine mL)')
plt.legend(loc = 2)
plt.suptitle(file_name, x = 0.1, y = -0.175, ha = 'left', size = 10)

plt.savefig(file_name + ' 3 Initial TC Concentration.png', dpi = 600, bbox_inches = 'tight')
plt.close()
# plt.show()

#%% ########## Export to CSV
table_position_unc = len(df_cal.index)

# Dataframe for primary export, including final urinary concentration (correcting for 10,000 fold dilution)

# Measured_TC_Conc is the concentration in the vial (the diluted form, ug creatinine / mL diuluted urine). 
df_spl_export = df_spl.loc[:, ['SampleID', 'IS_Recovery', 'Measured_TC_Conc']].sort_index()
df_spl_export['Batch'] = file_name
df_spl_export['Creatinine in Urine (ug creatinine / mL urine)'] = df_spl_export['Measured_TC_Conc'] * 10000
df_spl_export = df_spl_export.loc[:, ['Batch', 'SampleID', 'IS_Recovery', 'Creatinine in Urine (ug creatinine / mL urine)']]

# Compare corrected to initial quantification results and MassHunter results
df_spl_for_merge = df_spl.loc[:, ['SampleID', 'Measured_TC_Conc']]
df_comparison = pd.merge(df_spl_for_merge, df_spl_MassHunter, how='outer', on='SampleID')
df_comparison = df_comparison.rename(index=str, 
    columns={'Measured_TC_Conc':'Measured_TC_Conc_Uncorrected', 
    'Analyte_CalcConc':'Measured_TC_Conc_MassHunter'})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(file_name + '_Results' + '.xlsx', engine='xlsxwriter')
#df_spl_export = df_spl_export.rename(index=str, columns={'Measured_TC_Conc':analyte})
df_spl_export.to_excel(writer, sheet_name='Export')
df_comparison.to_excel(writer, sheet_name='Comparison')
df_cal.to_excel(writer, sheet_name='Accuracy')
df_spl.to_excel(writer, sheet_name='Accuracy', startrow=table_position_unc + 2)
writer.save()

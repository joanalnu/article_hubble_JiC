#IMPORANT THE DATA MUST BE DISPOSED AS FOLLOWING:
#ID column A; RAC and DEC columns B and C; PERIOD column D; V(mag) column E
#
#collect period data from excel file to compute absolute (visual) magnitude
#collect apparent magnitude to compute distance
import matplotlib.pyplot as plt
import xlwings as xw
import math
import numpy as np

def working(wb, sheet):
    if sheet['A1'].value == 'working':
        print('xlwings connection succesful!')

def aux_id(n):
    aux_id_data = list()
    for i in range(n):
        aux_id_data.append(i)
    return aux_id_data

#file_name = 'NGC_4258'
file_name = str(input('File name (without ".xlsx"): '))
file_path = '/Users/j.alcaide/Documents/cepheid/data/'+file_name+'.xlsx'
wb = xw.Book(file_path)
sheet = wb.sheets[0]
working(wb, sheet)
print('Using: ', file_path)

label_0 = sheet['A1'].value
label_1 = sheet['D1'].value

id_data = list()
cell = sheet['A2']
while cell.value:
    id_data.append(cell.value)
    cell = cell.offset(1, 0)

p_data = list()
cell = sheet['D2']
while cell.value:
    p_data.append(cell.value)
    cell = cell.offset(1, 0)



M_V_data = list()
M_V_error_up = list()
M_V_error_down = list()
for i in range(len(p_data)):
    M_V = -2.744*(math.log10(p_data[i])-1.4)-5.262
    M_V_data.append(M_V)
    M_V_up = -2.691*(math.log10(p_data[i])-1.4)-5.222
    M_V_error_up.append(abs(M_V_up-M_V))
    M_V_down = -2.857*(math.log10(p_data[i])-1.4)-5.302
    M_V_error_down.append(abs(M_V_down-M_V))
    
M_I_data = list()
M_I_error_up = list()
M_I_error_down = list()
for i in range(len(p_data)):
    M_I = -3.039*(math.log10(p_data[i])-1.4)-6.054
    M_I_data.append(M_I)
    M_I_up = -2.980*(math.log10(p_data[i])-1.4)-6.026
    M_I_error_up.append(abs(M_I_up-M_I))
    M_I_down = -3.098*(math.log10(p_data[i])-1.4)-6.082
    M_I_error_down.append(abs(M_I_down-M_I))


aux_id_data = aux_id(len(id_data))
fig1, ax = plt.subplots(2, 2, figsize=(12,13))
#cepheid vs magnitude
ax[0,0].scatter(aux_id_data, M_V_data, color='blue')
ax[0,0].set_xlabel('Cepheid ID')
ax[0,0].set_ylabel('absolute Magnitude M_V')
ax[0,0].errorbar(aux_id_data, M_V_data, yerr=[M_V_error_down, M_V_error_up], fmt='none', color='red', capsize=5, label='Error Bars')

ax[1,0].scatter(aux_id_data, M_I_data, color='orange')
ax[1,0].set_xlabel('Cepheid ID')
ax[1,0].set_ylabel('absolute Magnitude M_I')
ax[1,0].errorbar(aux_id_data, M_I_data, yerr=[M_I_error_down, M_I_error_up], fmt='none', color='red', capsize=5, label='Error Bars')

#period vs magnitudes
ax[0,1].scatter(p_data, M_V_data, color='blue')
ax[0,1].set_xlabel('Period [days]')
ax[0,1].set_ylabel('absolute Magnitude M_V')
ax[0,1].errorbar(p_data, M_V_data, yerr=[M_V_error_down, M_V_error_up], fmt='none', color='red', capsize=5, label='Error Bars')

ax[1,1].scatter(p_data, M_I_data, color='orange')
ax[1,1].set_xlabel('Period [days]')
ax[1,1].set_ylabel('absolute Magnitude M_I')
ax[1,1].errorbar(p_data, M_I_data, yerr=[M_I_error_down, M_I_error_up], fmt='none', color='red', capsize=5, label='Error Bars')



fig1_path = '/Users/j.alcaide/Documents/cepheid/'+file_name+'_Magnitudes_1.png'
fig1.savefig(fig1_path)
print('Figure 1 saved as: ', fig1_path)


print("END PART 1")
print("")



#THE FOLLOWING CODE IS ONLY CORRECT WHEN m = V(mag)
m_V_data = list()
cell = sheet['E2'] #Visual column is E
for i in range(len(id_data)):
    if cell.value == None:
        m_V_data.append(0)
    else:
        m_V_data.append(cell.value)
    cell = cell.offset(1,0)

m_I_data = list()
cell = sheet['F2']
for i in range(len(id_data)):
    if cell.value == None:
        m_I_data.append(None)
    else:
        m_I_data.append(cell.value)
    cell = cell.offset(1,0)

#DISTANCES
d_V_data = list()
d_V_error_up = list()
d_V_error_down = list()
for i in range(len(M_V_data)):
    if m_V_data[i] == 0:
        d_V_data.append(0)
    else:
        d_V = (10 ** ((m_V_data[i] - M_V_data[i])/5) )*10
        d_V_data.append(d_V)
        d_V_up = (10 ** ((m_V_data[i] - M_V_error_up[i])/5) )*10
        d_V_error_up.append(d_V_up)
        d_V_down = (10 ** ((m_V_data[i] - M_V_error_down[i])/5) )*10
        d_V_error_down.append(d_V_down)

d_I_data = list()#
d_I_error_up = list()#
d_I_error_down = list()#
for i in range(len(M_I_data)):#
    if m_I_data[i] == 0:#
        d_I_data.append(None)#
    else:#
        d_I = (10 ** ((m_I_data[i] - M_I_data[i])/5) )*10#
        d_I_data.append(d_I)#
        d_I_up = (10 ** ((m_I_data[i] - M_I_error_up[i])/5) )*10#
        d_I_error_up.append(d_I_up)#
        d_I_down = (10 ** ((m_I_data[i] - M_I_error_down[i])/5) )*10#
        d_I_error_down.append(d_I_down)#


#plotting distances and apparent magnitudes

fig2, axs = plt.subplots(2, 3, figsize=(12,13))

axs[0,0].scatter(m_V_data, M_V_data)
axs[0,0].set_xlabel('apparent magnitude m_V')
axs[0,0].set_ylabel('absolute magnitude M_V')
axs[0,0].set_title('correlation apparent to absolute magnitudes')
axs[0,0].errorbar(m_V_data, M_V_data, yerr=[M_V_error_down, M_V_error_up], fmt='none', color='gray', capsize=5, label='Error Bars')

axs[1,0].scatter(m_I_data, M_I_data)#
axs[1,0].set_ylabel('apparent magnitude m_I')#
axs[1,0].set_xlabel('apparent magnitude m_I')#
axs[1,0].errorbar(m_I_data, M_I_data, yerr=[M_V_error_down, M_V_error_up], fmt='none', color='gray', capsize=5, label='Error Bars')#

axs[0,1].scatter(m_V_data, d_V_data)
axs[0,1].set_xlabel('apparent magnitude m_V')
axs[0,1].set_ylabel('distance d_V in pc')
axs[0,1].set_title('distance per cepheid ID')
axs[0,1].errorbar(m_V_data, d_V_data, yerr=[d_V_error_down, d_V_error_up], fmt='none', color='gray', capsize=5, label='Error Bars')

axs[1,1].scatter(m_I_data, d_I_data)#
axs[1,1].set_xlabel('apparent magnitude m_I')#
axs[1,1].set_ylabel('distance d_I in pc')#
axs[1,1].errorbar(m_I_data, d_I_data, yerr=[d_I_error_down, d_I_error_up], fmt='none', color='gray', capsize=5, label='Error Bars')#


id_num = aux_id(len(id_data))

axs[0,2].scatter(id_num, d_V_data)
axs[0,2].tick_params(axis='x', rotation=45)
axs[0,2].set_xlabel('Cepheid ID')
axs[0,2].set_ylabel('distance d_V in pc')
axs[0,2].set_title('distance per cepheid ID')
axs[0,2].errorbar(id_num, d_V_data, yerr=[d_V_error_down, d_V_error_up], fmt='none', color='gray', capsize=5, label='Error Bars')
# best fit line to distance d_V
m, b = np.polyfit(id_num, d_V_data, 1)
best_fit_line_V = [m * id + b for id in id_num]
axs[0,2].plot(id_num, best_fit_line_V, label=f'Best Fit Line: y = {m:.2f}x + {b:.2f}', color='red')

axs[1,2].scatter(id_num, d_I_data)#
axs[1,2].tick_params(axis='x', rotation=45)#
axs[1,2].set_xlabel('Cepheid ID')#
axs[1,2].set_ylabel('distance d_I in pc')#
axs[1,2].errorbar(id_num, d_I_data, yerr=[d_I_error_down, d_I_error_up], fmt='none', color='gray', capsize=5, label='Error Bars')#
# best fit line to distance d_I
m, b = np.polyfit(id_num, d_I_data, 1)#
best_fit_line_I = [m * id + b for id in id_num]#
axs[1,2].plot(id_num, best_fit_line_I, label=f'Best Fit Line: y = {m:.2f}x + {b:.2f}', color='red')#


plt.legend()

fig_path_2 = '/Users/j.alcaide/Documents/cepheid/'+file_name+'_distances_1.png'
fig2.savefig(fig_path_2)
print('Figure 2 saved as: ', fig_path_2)
print("END PART 2")

print("")

#print magnitude medium
sum_V = 0
sum_I = 0
for i in range(len(id_data)):
    sum_V += M_V_data[i]
    sum_I += M_I_data[i]
print('mid V Magnitude', sum_V/len(M_V_data), end='\n')
print('mid I Magnitude', sum_I/len(M_I_data), end='\n')

#print distance medium
sum_V = 0
sum_I = 0
for i in range(len(id_data)):
    sum_V += d_V_data[i]
    sum_I += d_I_data[i]#

#cleaning before printing (excel file)
cell_1 = sheet['G1']
cell_2 = sheet['H1']
for i in range(10):
    cell_1.value = None
    cell_1.offset(1, 0)
    cell_2.value = None
    cell_2.offset(1,0)

#printing results
print('mid V distance', sum_V/len(d_V_data), ' pc', end='\n')
sheet['G2'].value = 'mid V distance [pc]'
sheet['H2'].value = (sum_V/len(d_V_data))
print('lowest: ', min(d_V_data), '  ;  highest: ', max(d_V_data))
sheet['G3'].value = 'lowest d [pc]'
sheet['H3'].value = min(d_V_data)
sheet['G4'].value = 'highest d [pc]'
sheet['H4'].value = max(d_V_data)
print('mid I distance', sum_I/len(d_I_data), ' pc', end='\n')#
sheet['G7'].value = 'mid I distance [pc]'
sheet['H7'].value = (sum_I/len(d_I_data))#
print('lowest: ', min(d_I_data), '  ;  highest: ', max(d_I_data))#
sheet['G8'].value = 'lowest d [pc]'
sheet['H8'].value = min(d_I_data)#
sheet['G9'].value = 'highest d [pc]'
sheet['H9'].value = max(d_I_data)#

with open('/Users/j.alcaide/Documents/cepheid/V_distances.txt', 'a') as file:
    file.write(file_name)
    file.write(' ')
    file.write(str((sum_V/len(d_V_data))))
    file.write(' ')
    file.write(str(min(d_V_data)))
    file.write(' ')
    file.write(str(min(d_V_data)))
    file.write('\n')

with open('/Users/j.alcaide/Documents/cepheid/I_distances.txt', 'a') as file:#
    file.write(file_name)#
    file.write(' ')#
    file.write(str((sum_I/len(d_I_data))))#
    file.write(' ')#
    file.write(str(min(d_I_data)))#
    file.write(' ')#
    file.write(str(min(d_I_data)))#
    file.write('\n')#


print("")
print('END')
print("")
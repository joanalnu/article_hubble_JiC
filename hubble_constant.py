import xlwings as xw
import matplotlib.pyplot as plt
import numpy as np
from scipy.odr import ODR, Model, Data

#reading redshift file
path_redshift = '/Users/j.alcaide/Documents/spectra/redshift.xlsx'
xw.Book(path_redshift).set_mock_caller()
wb = xw.Book.caller()
sheet = wb.sheets[0]

redshift_id = list()
cell = sheet['A1']
while cell.value:
    redshift_id.append(cell.value)
    cell = cell.offset(1, 0)

redshift_data = list()
cell = sheet['B1']
while cell.value:
    redshift_data.append(cell.value)
    cell = cell.offset(1, 0)

aux_id = list()
for i in range(len(redshift_id)):
    aux_id.append(i+1)

redshift_error = list()
cell = sheet['C1']
while cell.value:
    redshift_error.append(cell.value)
    cell = cell.offset(1, 0)

redshift_error_down = list()
redshift_error_up = list()
for i in range(len(redshift_error)):
    redshift_error_down.append(abs(redshift_data[i] - redshift_error[i]))
    redshift_error_up.append(abs(redshift_data[i] + redshift_error[i]))

#for sorting
sorting_redshift_lists = list(zip(redshift_id, redshift_data, redshift_error_down, redshift_error_up))
sorted_redshift_lists = sorted(sorting_redshift_lists, key=lambda x: x[1])
sorted_redshift_id, sorted_redshift_data, sorted_redshift_error_down, sorted_redshift_error_up = zip(*sorted_redshift_lists)

velocities_id = redshift_id
velocities_data = list()
velocities_error_down = list()
velocities_error_up = list()
for i in range(len(redshift_id)):
    velocities_data.append(redshift_data[i]*299792458)
    velocities_error_down.append(redshift_error[i]*299792458) #abs(()*)
    velocities_error_up.append(redshift_error[i]*299792458)

#for sorting
sorting_velocities_lists = list(zip(velocities_id, velocities_data, velocities_error_down, velocities_error_up))
sorted_velocities_lists = sorted(sorting_velocities_lists, key=lambda x: x[1])
sorted_velocities_id, sorted_velocities_data, sorted_velocities_error_down, sorted_velocities_error_up = zip(*sorted_velocities_lists)

fig1, axs = plt.subplots(2, 2, figsize=(12, 13))
axs[0,0].scatter(redshift_id, redshift_data)
axs[0,0].set_xlabel('Galaxy ID')
axs[0,0].tick_params(axis='x', rotation=45, labelsize=5)
axs[0,0].set_ylabel('Redshift')
axs[0,0].set_title('Galaxy Redshift')
axs[0,0].grid(True)
axs[0,0].errorbar(redshift_id, redshift_data, yerr=[redshift_error_down, redshift_error_up], fmt='none', color='gray', capsize=5)

axs[0,1].scatter(sorted_redshift_id, sorted_redshift_data)
axs[0,1].set_xlabel('Galaxy ID')
axs[0,1].tick_params(axis='x', rotation=45, labelsize=5)
axs[0,1].set_ylabel('Redshift')
axs[0,1].set_title("Galaxy Redshift sorted")
axs[0,1].grid(True)
axs[0,1].errorbar(sorted_redshift_id, sorted_redshift_data, yerr=[sorted_redshift_error_down, sorted_redshift_error_up], fmt='none', color='gray', capsize=5)

axs[1,0].scatter(velocities_id, velocities_data)
axs[1,0].set_xlabel('Galaxy ID')
axs[1,0].tick_params(axis='x', rotation=45, labelsize=5)
axs[1,0].set_ylabel('Redshift')
axs[1,0].set_title("Galaxy velocity")
axs[1,0].grid(True)
axs[1,0].errorbar(velocities_id, velocities_data, yerr=[velocities_error_down, velocities_error_up], fmt='none', color='gray', capsize=5)

axs[1,1].scatter(sorted_velocities_id, sorted_velocities_data)
axs[1,1].set_xlabel('Galaxy ID')
axs[1,1].tick_params(axis='x', rotation=45, labelsize=5)
axs[1,1].set_ylabel('Redshift')
axs[1,1].set_title("Galaxy velocity sorted")
axs[1,1].grid(True)
axs[1,1].errorbar(sorted_velocities_id, sorted_velocities_data, yerr=[sorted_velocities_error_down, sorted_velocities_error_up], fmt='none', color='gray', capsize=5)

path_velocities_fig = '/Users/j.alcaide/Desktop/velocities_fig1.png'
fig1.savefig(path_velocities_fig)



V_gal_id = list()
V_dis_data = list()
V_dis_low = list()
V_dis_high = list()
graphic_V_dis_low = list()
graphic_V_dis_high = list()
with open('/Users/j.alcaide/Documents/cepheid/V_distances.txt','r') as file:
    lines = file.readlines()

for line in lines:
    data = line.split()
    if len(data) == 4:
        V_gal_id.append(data[0])
        V_dis_data.append(float(data[1]))
        V_dis_low.append(float(data[1])-float(data[2]))
        graphic_V_dis_low.append(float(data[2]))
        V_dis_high.append(float(data[1])-float(data[3]))
        graphic_V_dis_high.append(float(data[3]))
    else:
        continue


I_gal_id = list()
I_dis_data = list()
I_dis_low = list()
I_dis_high = list()
graphic_I_dis_low = list()
graphic_I_dis_high = list()
with open('/Users/j.alcaide/Documents/cepheid/I_distances.txt','r') as file:
    lines = file.readlines()

for line in lines: 
    data = line.split()
    if len(data) == 4:
        I_gal_id.append(data[0])
        I_dis_data.append(float(data[1]))
        I_dis_low.append(float(data[1])-float(data[2]))
        graphic_I_dis_low.append(float(data[2]))
        I_dis_high.append(float(data[1])-float(data[3]))
        graphic_I_dis_high.append(float(data[3]))
    else:
        continue


fig2, axs = plt.subplots(1, 2, figsize=(12,13))
axs[0].scatter(V_gal_id, V_dis_data)
axs[0].set_xlabel('Galaxy ID')
axs[0].tick_params(axis='x', rotation=45, labelsize=5)
axs[0].set_ylabel('V_distance')
axs[0].set_title('Galaxy distances V')
axs[0].grid(True)
axs[0].errorbar(V_gal_id, V_dis_data, yerr=[V_dis_low, V_dis_high], fmt='none', color='gray', capsize=5)

axs[1].scatter(I_gal_id, I_dis_data)
axs[1].set_xlabel('Galaxy ID')
axs[1].tick_params(axis='x', rotation=45, labelsize=5)
axs[1].set_ylabel('I_distance')
axs[1].set_title("Galaxy distances I")
axs[1].grid(True)
axs[1].errorbar(I_gal_id, I_dis_data, yerr=[I_dis_low, I_dis_high], fmt='none', color='gray', capsize=5)

path_distances_fig = '/Users/j.alcaide/Desktop/distances_fig2.png'
fig2.savefig(path_distances_fig)


#computing the Hubble Constant 😁

hubble_combined_lists = list(zip(velocities_id, I_gal_id, velocities_data, velocities_error_down, velocities_error_up, V_dis_data, graphic_V_dis_low, graphic_V_dis_high, I_dis_data, graphic_I_dis_low, graphic_I_dis_high))
sorted_hubble_lists = sorted(hubble_combined_lists, key=lambda x: x[5])
sorted_galaxy_id, unuseful_list, sorted_vel_data, sorted_vel_low, sorted_vel_high, sorted_V_dis_data, sorted_V_dis_low, sorted_V_dis_high, sorted_I_dis_data, sorted_I_dis_low, sorted_I_dis_high = zip(*sorted_hubble_lists)

#convert pc -> Mpc and m/s -> KM/s
converted_sorted_vel_data = list()
converted_sorted_vel_low = list()
converted_sorted_vel_high = list()
for i in range(len(sorted_vel_data)):
    converted_sorted_vel_data.append(sorted_vel_data[i] / 1000) #m/s -> km/s
    converted_sorted_vel_low.append(sorted_vel_low[i] / 1000)
    converted_sorted_vel_high.append(sorted_vel_high[i] / 1000)

converted_sorted_V_dis_data = list()
converted_sorted_V_dis_low = list()
converted_sorted_V_dis_high = list()
for i in range(len(sorted_V_dis_data)):
    converted_sorted_V_dis_data.append(sorted_V_dis_data[i] / 1000000) #pc -> Mpc
    converted_sorted = abs(sorted_V_dis_low[i] / 1000000)
    converted_sorted_V_dis_low.append(converted_sorted_V_dis_data[i] - converted_sorted)
    converted_sorted = abs(sorted_V_dis_high[i] / 1000000)
    converted_sorted_V_dis_high.append(converted_sorted_V_dis_data[i] - converted_sorted)

converted_sorted_I_dis_data = list()
converted_sorted_I_dis_low = list()
converted_sorted_I_dis_high = list()
for i in range(len(sorted_I_dis_data)):
    converted_sorted_I_dis_data.append(sorted_I_dis_data[i] / 1000000) #pc -> Mpc
    converted_sorted = abs(sorted_I_dis_low[i] / 1000000)
    converted_sorted_I_dis_low.append(converted_sorted_I_dis_data[i] - converted_sorted)
    converted_sorted = abs(sorted_I_dis_high[i] / 1000000)
    converted_sorted_I_dis_high.append(converted_sorted_I_dis_data[i] - converted_sorted)


fig3, ax = plt.subplots(1, 1, figsize=(18, 20))
ax.scatter(converted_sorted_V_dis_data, converted_sorted_vel_data, marker='x')
ax.set_xlabel('Distances [Mpc]')
ax.set_ylabel('Velocities km/s')
ax.set_title('Hubble Diagram')
ax.grid(True)

ax.errorbar(converted_sorted_V_dis_data, converted_sorted_vel_data, yerr=[converted_sorted_vel_low, converted_sorted_vel_high], fmt='none', color='red', capsize=5)
ax.errorbar(converted_sorted_V_dis_data, converted_sorted_vel_data, xerr=[converted_sorted_V_dis_low, converted_sorted_V_dis_high], fmt='none', color='blue', capsize=5)


ax.scatter(converted_sorted_I_dis_data, converted_sorted_vel_data, color='green', marker='x')
ax.errorbar(converted_sorted_I_dis_data, converted_sorted_vel_data, yerr=[converted_sorted_vel_low, converted_sorted_vel_high], fmt='none', color='red', capsize=5)
ax.errorbar(converted_sorted_I_dis_data, converted_sorted_vel_data, xerr=[converted_sorted_I_dis_low, converted_sorted_I_dis_high], fmt='none', color='green', capsize=5)

converted_sorted_average_dis_data = list() #average between V and I
converted_sorted_average_dis_low = list()
converted_sorted_average_dis_high = list()
for i in range(len(converted_sorted_I_dis_data)):
    converted_sorted_average_dis_data.append((converted_sorted_V_dis_data[i]+converted_sorted_I_dis_data[i])/2)
    converted_sorted_average_dis_low.append((converted_sorted_V_dis_low[i]+converted_sorted_I_dis_low[i])/2)
    converted_sorted_average_dis_high.append((converted_sorted_V_dis_high[i]+converted_sorted_I_dis_high[i])/2)

# compute fit
x = np.array(converted_sorted_average_dis_data)

converted_sorted_average_dis_low_combined = list()
for i in range(len(converted_sorted_average_dis_low)):
    converted_sorted_average_dis_low_combined.append(converted_sorted_average_dis_data[i]+converted_sorted_average_dis_low[i])
x_low = np.array(converted_sorted_average_dis_low_combined)

converted_sorted_average_dis_high_combined = list()
for i in range(len(converted_sorted_average_dis_high)):
    converted_sorted_average_dis_high_combined.append(converted_sorted_average_dis_data[i]+converted_sorted_average_dis_high[i])
x_high = np.array(converted_sorted_average_dis_high_combined)

y = np.array(converted_sorted_vel_data)
y_low = np.array(converted_sorted_vel_low)
y_high = np.array(converted_sorted_vel_high)

# Define the linear model with scripy ODR
def linear_func(B, x):
    return B[0] * x + B[1]

data = Data(x, y) #, wd=1/np.array(distances_error))
model = Model(linear_func)

# run fitting
odr = ODR(data, model, beta0=[1, 0])
output = odr.run()

# Extract results
H0 = output.beta[0]
H0_err = output.sd_beta[0]
intercept = output.beta[1]
print(f'Hubble constant = {H0:.2f} ± {H0_err:.2f}')

fit_linear_regression = H0 * np.array(x) + intercept
fit_low = (H0-H0_err) * np.array(x) + intercept
fit_high = (H0+H0_err) * np.array(x) + intercept






ax.plot(x, fit_linear_regression, color='red')

residuals = y - fit_linear_regression
weighted_residuals = residuals / converted_sorted_average_dis_low
mse = np.sum(weighted_residuals ** 2) / (len(x) - 2)
variance = np.sum((x - np.mean(x)) ** 2)
std_error_slope = np.sqrt(mse / variance)





xmin, xmax = plt.xlim()
ymin, ymax = plt.ylim()

plt.text(xmax * 0.95, ymax * 0.98, fontsize=20, ha='right', va='top')

fig3.savefig('/Users/j.alcaide/Desktop/hubble_diagram.png', dpi=500, bbox_inches='tight')


#####################################################################################################################


aux_id = list()

def function_aux_id(galaxy_names):
    aux_id.clear()
    for i in range(len(galaxy_names)):
        aux_id.append(i)

# plots for the paper
        
fig_hubble, ax = plt.subplots(1, 1, figsize=(10, 12))
ax.scatter(converted_sorted_average_dis_data, converted_sorted_vel_data, marker='x')
ax.set_xlabel('Distances [Mpc]')
ax.set_ylabel('Velocities [km/s]')
ax.grid(True)

ax.errorbar(converted_sorted_average_dis_data, converted_sorted_vel_data, yerr=[converted_sorted_vel_low, converted_sorted_vel_high], fmt='none', color='gray', capsize=5)
ax.errorbar(converted_sorted_average_dis_data, converted_sorted_vel_data, xerr=[converted_sorted_average_dis_low, converted_sorted_average_dis_high], fmt='none', color='gray', capsize=5)

ax.plot(x, fit_linear_regression, color='red')

fig_hubble.savefig('/Users/j.alcaide/Desktop/fig_hubble.png', dpi=500, bbox_inches='tight')

for i in range(len(aux_id)):
    aux_id[i] += 1

#distances
    
for i in range(len(V_dis_data)):
    V_dis_data[i] = V_dis_data[i] / 1000
    V_dis_low[i] = V_dis_low[i] / 1000
    V_dis_high[i] = V_dis_high[i] / 1000

    I_dis_data[i] = I_dis_data[i] / 1000
    I_dis_low[i] = I_dis_low[i] / 1000
    I_dis_high[i] = I_dis_high[i] / 1000


function_aux_id(V_gal_id)
fig_V_distances, axs = plt.subplots(1, 1, figsize=(10,9))
axs.scatter(aux_id, V_dis_data)
axs.set_xlabel('Galaxy ID')
axs.set_ylabel('distance V band [kpc]')
axs.grid(False)
axs.errorbar(aux_id, V_dis_data, yerr=[V_dis_low, V_dis_high], fmt='none', color='gray', capsize=5)

fig_V_distances.savefig('/Users/j.alcaide/Desktop/fig_V_distances.png', dpi=300, bbox_inches='tight')


fig_I_distances, axs = plt.subplots(1, 1, figsize=(10,9))
function_aux_id(I_gal_id)
axs.scatter(aux_id, I_dis_data)
axs.set_xlabel('Galaxy ID')
axs.set_ylabel('distance I band [kpc]')
axs.grid(False)
axs.errorbar(aux_id, I_dis_data, yerr=[I_dis_low, I_dis_high], fmt='none', color='gray', capsize=5)

fig_I_distances.savefig('/Users/j.alcaide/Desktop/fig_I_distances.png', dpi=300, bbox_inches='tight')



#redshifts and velocities

this_redshift_data = list()
this_redshift_error_up = list()
this_redshift_error_down = list()
this_velocities_data = list()
this_velocities_error_up = list()
this_velocities_error_down = list()
aux_id.clear()
for i in range(49):
    aux_id.append(i)
    this_redshift_data.append(redshift_data[i])
    this_redshift_error_up.append(redshift_error_up[i])
    this_redshift_error_down.append(redshift_error_down[i])
    this_velocities_data.append(velocities_data[i])
    this_velocities_error_up.append(velocities_error_up[i])
    this_velocities_error_down.append(velocities_error_down[i])

fig_redshifts, axs = plt.subplots(1, 1, figsize=(7,7))
axs.scatter(aux_id, this_redshift_data)
axs.set_xlabel('Galaxy ID')
axs.set_ylabel('Redshift')
axs.grid(True)
axs.errorbar(aux_id, this_redshift_data, yerr=[this_redshift_error_down, this_redshift_error_up], fmt='none', color='gray', capsize=5)
plt.xticks(ticks=np.arange((len(aux_id))), fontsize=4)
fig_redshifts.savefig('/Users/j.alcaide/Desktop/fig_redshifts.png', dpi=200, bbox_inches='tight')

for i in range(len(aux_id)):
    aux_id[i]+=1


fig_velocities, axs = plt.subplots(2, 1,  gridspec_kw={'height_ratios': [1, 3]})

for i in range(len(this_velocities_data)):
    this_velocities_data[i] = this_velocities_data[i] / 1000
    this_velocities_error_down[i] = this_velocities_error_down[i] / 1000
    this_velocities_error_up[i] = this_velocities_error_up[i] / 1000

axs[0].scatter(aux_id, this_velocities_data, color="tab:blue")
axs[0].set_xlabel('Galaxy ID')
axs[0].set_ylabel('velocities [km/s]  ')
axs[0].errorbar(aux_id, this_velocities_data, yerr=[this_velocities_error_down, this_velocities_error_up], fmt='none', color='gray', capsize=5)
axs[0].set_ylim(bottom=2600, top=None, emit=True, auto=False, ymin=None, ymax=None)

axs[1].scatter(aux_id, this_velocities_data, color="tab:blue")
axs[1].set_xlabel('Galaxy ID')
axs[1].set_ylabel('velocities [km/s]')
axs[1].errorbar(aux_id, this_velocities_data, yerr=[this_velocities_error_down, this_velocities_error_up], fmt='none', color='gray', capsize=5)
axs[1].set_ylim(bottom=-500, top=2600, emit=True, auto=False, ymin=None, ymax=None)


fig_velocities.tight_layout()
fig_velocities.savefig('/Users/j.alcaide/Desktop/fig_velocities.png', dpi=200, bbox_inches='tight')


path_redshift = '/Users/j.alcaide/Documents/Table_galaxy_sample_overleaf.xlsx'
xw.Book(path_redshift).set_mock_caller()
wb = xw.Book.caller()
sheet = wb.sheets[0]

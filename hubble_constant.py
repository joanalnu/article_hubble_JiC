# importing required libraries
import xlwings as xw # TODO: migrate data to a csv file
import matplotlib.pyplot as plt
import numpy as np
from scipy.odr import ODR, Model, Data

# reading data
# redshift
path_redshift = './redshift_data/00_redshift.xlsx'
xw.Book(path_redshift).set_mock_caller()
wb = xw.Book.caller()
sheet = wb.sheets[0]

redshift_id, redshift_data, redshift_error, aux_id = [], [], [], []
id_cell, data_cell, error_cell = sheet['A1'], sheet['B1'], sheet['C1']
while id_cell.value:
    redshift_id.append(id_cell.value)
    redshift_data.append(data_cell.value)
    redshift_error.append(error_cell.value)
    id_cell, data_cell, error_cell = id_cell.offset(1, 0), data_cell.offset(1, 0), error_cell.offset(1, 0)
    aux_id.append(len(redshift_id))

velocities_id = redshift_id
velocities_data = [data * 299792458 for data in redshift_data]
velocities_error = [data *299792458 for data in redshift_error]

fig1, ax = plt.subplots(1, 2) #, figsize=(12, 13))
ax[0].scatter(redshift_id, redshift_data)
ax[0].set_xlabel('Galaxy ID')
ax[0].tick_params(axis='x', rotation=45, labelsize=5)
ax[0].set_ylabel('Redshift')
ax[0].set_title('Galaxy Redshift')
ax[0].grid(True)
ax[0].errorbar(redshift_id, redshift_data, yerr=redshift_error, fmt='none', color='gray', capsize=5)

ax[1].scatter(velocities_id, velocities_data)
ax[1].set_xlabel('Galaxy ID')
ax[1].tick_params(axis='x', rotation=45, labelsize=5)
ax[1].set_ylabel('Redshift')
ax[1].set_title("Galaxy velocity")
ax[1].grid(True)
ax[1].errorbar(velocities_id, velocities_data, yerr=velocities_error, fmt='none', color='gray', capsize=5)

fig1.savefig('./figs/velocities.png', dpi=600)

# distances
v_gal_id, v_data, v_error = [], [], []
with open('./cepheid_data/00_V_distances.txt', 'r') as f:
    lines = f.readlines()

for line in lines:
    data = line.split()
    if len(data) == 4:
        v_gal_id.append(data[0])
        v_data.append(float(data[1]))
        error = float(data[1]) - float(data[2])
        v_error.append(error)

i_gal_id, i_data, i_error = [], [], []
with open ('./cepheid_data/00_I_distances.txt', 'r') as f:
    lines = f.readlines()

for line in lines:
    data = line.split()
    if len(data)==4:
        i_gal_id.append(data[0])
        i_data.append(float(data[1]))
        error = float(data[1]) - float(data[2])
        i_error.append(error)


fig2, ax = plt.subplots(1, 2)
ax[0].scatter(v_gal_id, v_data)
ax[0].set_xlabel('Galaxy ID')
ax[0].tick_params(axis='x', rotation=45, labelsize=5)
ax[0].set_ylabel('V_distance')
ax[0].set_title('Galaxy distances V')
ax[0].grid(True)
ax[0].errorbar(v_gal_id, v_data, yerr=v_error, fmt='none', color='gray', capsize=5)

ax[1].scatter(i_gal_id, i_data)
ax[1].set_xlabel('Galaxy ID')
ax[1].tick_params(axis='x', rotation=45, labelsize=5)
ax[1].set_ylabel('I_distance')
ax[1].set_title("Galaxy distances I")
ax[1].grid(True)
ax[1].errorbar(i_gal_id, i_data, yerr=i_error, fmt='none', color='gray', capsize=5)

fig2.savefig('./figs/distances.png', dpi=600)


# computing the Hubble constant

# m/s -> km/s
velocities_data = [data/1000 for data in velocities_data]
velocities_error = [data/1000 for data in velocities_error]
# pc -> Mpc
v_data = [data/1000000 for data in v_data]
v_error = [data/1000000 for data in v_error]
i_data = [data/1000000 for data in i_data]
i_error = [data/1000000 for data in i_error]


# starting figure
fig3, ax = plt.subplots(1, 1, figsize=(8, 10))

ax.scatter(v_data, velocities_data, color='blue')
ax.scatter(i_data, velocities_data, color='green')

ax.errorbar(v_data, velocities_data, yerr=velocities_error, fmt='none', color='gray', capsize=5)
ax.errorbar(v_data, velocities_data, xerr=v_error, fmt='none', color='gray', capsize=5)

ax.errorbar(i_data, velocities_data, yerr=velocities_error, fmt='none', color='gray', capsize=5)
ax.errorbar(i_data, velocities_data, yerr=i_error, fmt='none', color='gray', capsize=5)

# computing mean for both filters
distances = list()
distances_error = list()
for i in range(len(v_data)):
    distances.append((v_data[i]+i_data[i])/2)
    distances_error.append((v_error[i]+i_error[i])/2)

# compute fit
# Define the linear model
def linear_func(B, x):
    return B[0] * x + B[1]  # B[0] = slope, B[1] = intercept
model = Model(linear_func)

# taking into account distance errorbars
data = Data(distances, velocities_data)

# Run fitting model
odr = ODR(data, model, beta0=[1, 0])  # Initial guess: slope=1, intercept=0
output = odr.run()

# Extract results
H0 = output.beta[0]
H0_err = output.sd_beta[0]
intercept = output.beta[1]

# Print results
print(f'H0 = {H0:.4f} ± {H0_err:.4f}')

fit_function = H0 * np.array(distances) + intercept
fit_low = (H0-H0_err) * np.array(distances) + intercept
fit_high = (H0+H0_err) * np.array(distances) + intercept

ax.plot(distances, fit_function, color='r')
ax.fill_between(distances, fit_low, fit_high, color='orange', alpha=0.3)

plt.title("Hubble Diagram")
plt.xlabel("$d$ in $Mpc$")
plt.ylabel("$v$ in $km s^{-1}$")
plt.tight_layout()
plt.grid('True')
fig3.savefig('./figs/hubblediagram.png', dpi=600)




# Hubble constant with inverse polyfit
# This confirms the axis-bias in using the `np.polyfit()` function.

# x = np.array(velocities_data)
# y = np.array(distances)
#
# slope, intercept = np.polyfit(x, y, 1)
# fit_linear_regression = (1/slope) * x + intercept
# residuals = x - fit_linear_regression
# weighted_residuals = residuals / np.array(distances_error)
# mse = np.sum(weighted_residuals ** 2) / (len(x) - 2)
# variance = np.sum((x - np.mean(x)) ** 2)
# std_error_slope = np.sqrt(mse / variance)
# print(f'Hubble Constant: {1/slope} ± {std_error_slope}')

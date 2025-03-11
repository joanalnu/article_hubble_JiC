# article_hubble_JiC
This are the python codes used for the paper "Recomputing the Hubble constant", which is written as part of the Youth and Science program by Fundació Catalunya La Pedrera.

In summer 2024, I discovered `np.polyfit` prioritizes minimizing the vertical distance to data points, which can lead to wrong/unprecise values. Using `scipy.ord` module avoids this error, as it accounts equally for both axis (x and y directions), thus usually preferred for scientific curve fitting. This is something I wasn't aware when I conducted this project in 2023 (11th grade), as this was the first time I was using python for research and my first big project. I'll use `scipy` from now on ;)

Using `scipy` instead of `numpy` my data results in a value for H0 of $70.26±5.07 km s^{-1} Mpc^{-1}$. When plotting the inverse Hubble Diagram (distance in x-axis and velocities in y-axis) and computing H0 with `numpy` it results in a similar value, confirming the axis-bias when using `polyfit`.

If you consult the research paper and the code right now, it still is not updated regarding this error. I've corrected the code (which you can see in the `dev` directory in the `fit-function` branch), and am currently working on updating the paper.
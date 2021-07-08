# ecomonetrics_VBA

# Goldfield-Quandt_VBA
makro for Excel for Goldfield-Quandt heteroscedasticity test

y-data must be first, sort by yourself, select data and start macro

like this


<img src = "https://github.com/Dranikf/ecomonetrics_VBA/blob/main/examples/goldfild_example.JPG" height = "800" width = 600>

this example can be founded if examples folder

# reshape fuction

if you have range with NxM shape you can create a new one with shape KxL (N*M=K*L).
it will take values line by line and put them by lines. simple examples in "examples" file

# forecast_error funciton

funciton realise formula
<img src="https://render.githubusercontent.com/render/math?math=s_{f}%20=s\sqrt{%201%20%2B%20\frac{1}{n}%20%2B%20\frac{(x_f-\bar{x})}{n*\sigma_{x}^{2}}}">

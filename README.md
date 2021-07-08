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
<img src="https://render.githubusercontent.com/render/math?math=s_{f}%20=s\sqrt{%201%20%2B%20\frac{1}{n}%20%2B%20\frac{(x_f-\bar{x})}{n\sigma_{x}^{2}}}">
which ofen used for estimation of forecast error in regression models. It takes x value for forecating, range with x observations and a standart error of regression model. It returns value of given formula.

# basic_autocorrelation function

it automates the calculation of the autocorrelations of a series, by type of "shift" and use  "Correl" Excel funtion.
function takes range with target series and lag. For example:

<img src="https://github.com/Dranikf/ecomonetrics_VBA/blob/main/examples/correlogramm.jpg">

# QstatLB function

Computes Ljung–Box statics (https://en.wikipedia.org/wiki/Ljung%E2%80%93Box_test).
If computing statiscs for lag t, you should give you must pass the autocorrelation coefficients from 1st to tth to the function
and the size of a sample. For example:

<img src="https://github.com/Dranikf/ecomonetrics_VBA/blob/main/examples/Q%20stats%20funciton.JPG">
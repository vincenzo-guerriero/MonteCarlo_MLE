# MonteCarlo_MLE
This spreadsheet performs the Monte Carlo simulations illustrated in the paper entitled "Maximum Likelihood instead of Least Squares in fracture analysis by means of a simple Excel sheet with VBA macro", using the command [ctrl]+[m]. Nevertheless, it can be used to analyze your own field data. Don't forget to allow Excel Macros when opening the file. Furthermore, Excel's language settings must be set to "English", otherwise the formulas in the cells will not be interpreted correctly.

Routines in this spreadsheet:
Sub Initialize_Apert_Data(): initializes aperture data by filling columns F and G with opportune aperture class limits
Sub Maximize() and Sub Minimize_LS(): analyse a data set by maximizing or minimizing an object function, which is Log likelihood for the former and residual standard deviation for the latter.
Sub Simul_Apert_Data(): based on field data and on model true values of m, n and d, it produces a 35 entries data set by (i) resampling thickness data and (ii) producing, for each thickness, a random aperture value. Then, it identifies which class it belongs to and its limits.
Sub Simul_100(): For each triplet of true values m, n and d, it produces 100 simulated data sets and analyses each one by means of Sub Maximize(), in sheet MLE, and Sub Minimize_LS(), in sheet LSM. Then it saves the estimated values in columns S, T and U.
Sub Monte_Carlo(): varies the n true value in the range 0.05-4.75 mm, and for each value produces simulations by calling Sub Simul_100(), then saves the related results in sheet Results.

This spreadsheet can be used to analyze your own data. The fracture thickness and opening data will be entered in columns C and E respectively. Then call the routine Sub Initialize_Apert_Data(), which will insert the appropriate values into columns F and G. Then, call the Sub Maximize() on the MLE sheet, which will maximize the object function. This routine only works on that sheet. 
The sheet, in its current form, analyzes a data set of 35 items. When analyzing a data set of different size, by way of example 50 items, you should update the arrays in columns I, J, K, M, N, O and P so that they have the same size (i.e. 50). Furthermore, the formulas in cells J55, J46, K46 and L46 must be updated so that they perform calculations on arrays of suitable length.
The sheet also allows you to study the distribution of the residuals (deviations with respect to the best-fit line) in order to draw probability plots. To do this, simply copy the array of residuals in column J and paste it (with the "values only" option) into column N, then sort in ascending order.
Diagrams have been removed as they slow down the execution of routines.


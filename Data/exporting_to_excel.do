
* Load the dataset 
clear all
use "2024.dta"

export excel using "2024.xlsx", replace first(varl) locale(C)

// and then repeat the same for the other .dta files you want 


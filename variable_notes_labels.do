* ************************************************************************
* https://github.com/evpu
* Export variable names, labels and notes to excel
* ************************************************************************

clear all

* Open you file with data
use data

* select all variables and store the list in a local
ds *
local var_list `r(varlist)'
local var_number: list sizeof var_list


forval i = 1 / `var_number' {
    local var : word `i' of `var_list'
    * save variable label and notes (assume max is 2 notes per variable - modify accordingly if more needed)
    local label_`i' : variable label `var'
    notes _fetch note1_`i' : `var' 1
    notes _fetch note2_`i' : `var' 2
}


* Create a file to save all this info
clear
gen variable = ""
gen label = ""
gen note1 = ""
gen note2 = ""

set obs `var_number'

forval i = 1 / `var_number' {
    local var : word `i' of `var_list'
    replace variable = "`var'" in `i'
    replace label = "`label_`i''" in `i'
    replace note1 = `"`note1_`i''"' in `i'
    replace note2 = `"`note2_`i''"' in `i'
}

export excel "codebook.xlsx", sheet("Variables") sheetreplace firstrow(var)

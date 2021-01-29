* Descriptive stats & association with maltreatment alone



*Creating maltreatment score based on how many different exposures to maltreatment 

gen mal_score =  physical_abuse + sexual_abuse + emotional_abuse + emotional_neglect + physical_neglect

label var mal_score "Maltreatment score"


*Rename variables
*Match the names used in the other analysis do files
rename drinks_per_week alcohol 
rename csi smoking 
rename t2_inc diabetes_t2
rename CHD_inc chd 
rename af_inc af
rename Stroke_inc stroke 
rename ldl_c ldl

*Specify global macro for all phenotypic/outcome varibales 
global pheno_risk bmi alcohol ldl smoking sbp af diabetes_t2 chd stroke

*Specify global macro phenotypes according to whether continuous or binary 
global cont_pheno bmi alcohol ldl smoking sbp
global bin_pheno af diabetes_t2 chd stroke

*Standardise continuous variables for linear regressions

foreach var of varlist $cont_pheno {

	egen sd_`var'= std(`var')
	
}

*Take log of standardised continuous variables for regressions on multiplicative scale
*need to do (x+1) for lifetime smoking and drinks per week to avoid missing values 
foreach var of varlist smoking alcohol {

gen a_`var' = `var' + 1

}

foreach var of varlist bmi sbp a_smoking a_alcohol ldl {

gen ln_`var' = ln(`var')

}

*Standardise the logs we have just taken from continuous variables
foreach var of varlist ln_bmi ln_sbp ln_a_smoking ln_a_alcohol ln_ldl {

	egen sd_`var'= std(`var')
	
}

**********Touse 
*Exclude from sample individuals without maltreatment or GRS data 
mark touse
markout touse mal_score grs_*
keep if touse == 1 

*Create touse variable for each of the phenotypes 
foreach var of varlist $pheno_risk {

	mark touse_`var'
	markout touse_`var' mal_score age sex `var'
}

**************************** Descriptive stats *********************************

*As there as so few values to store I thought it might be possible to store everything under the same sheet 
putexcel set filename, sheet(all) modify
putexcel A1="Variable" B1="N" C1="Mean" D1="SD" E1="Controls" F1="Control (Percent)" G1="Cases" H1="Case (Percent)"

**********Continuous variables 

*Set up macro for all the continous variables- including age and sex 
global cont_stats $cont_pheno age number_sibling
local x = 1

foreach pheno of varlist $cont_stats {

	*specify the value of X to specify the excel cells to use 
	local x = `x' + 1

	*call the descriptive stats file
	putexcel set filename, sheet(all) modify

	*use summarise command to get SD, mean and sample number
	sum `pheno'
	
	*s
	local n_sample = r(N)
	local sd = r(sd)
	local n_mean = r(mean)
	local name : var label `pheno'

	putexcel A`x'="`name'" B`x'=`n_sample' C`x'=`n_mean' D`x'=`sd'

}

***********Binary variables 

*set the cell count to the number of the last cell 
*set up local macro of all binary variables - including maternal smoking 
global bin_stats $bin_pheno sex maternal_smoke

local x = 8
foreach pheno of varlist $bin_stats {

	*specify the value of X to specify the excel cells to use 
	local x = `x' + 1

	*call the descriptive stats file
	putexcel set filename, sheet(all) modify

	*tab binary outcomes with the matcell option so that the frequency values are stored in a matrix called results
	tab `pheno', matcell(results)
	
	*Pull out relevant results from the matrix created above to store in the excel file 
	local n_sample = r(N)
	local cases = results[2,1]
	local controls = results[1,1]
	local name : var label `pheno'

	putexcel A`x'="`name'" B`x'=`n_sample' E`x'=`controls' G`x'=`cases' 

}

******************** Entire Biobank stats **************************************

*Open file with data for all of Biobank's participants
use "filename", clear 


*Creating maltreatment score based on how many different exposures to maltreatment 

gen mal_score =  physical_abuse + sexual_abuse + emotional_abuse + emotional_neglect + physical_neglect

label var mal_score "Maltreatment score"


*Rename variables
*Match the names used in the other analysis do files
rename drinks_per_week alcohol 
rename csi smoking 
rename t2_inc diabetes_t2
rename CHD_inc chd 
rename af_inc af
rename Stroke_inc stroke 
rename ldl_c ldl

*Specify global macro for all phenotypic/outcome varibales 
global pheno_risk bmi alcohol ldl smoking sbp af diabetes_t2 chd stroke

*Specify global macro phenotypes according to whether continuous or binary 
global cont_pheno bmi alcohol ldl smoking sbp
global bin_pheno af diabetes_t2 chd stroke


**********Continuous variables 

*Set up macro for all the continous variables- including age and sex 
global cont_stats $cont_pheno age number_sibling
local x = 15

foreach pheno of varlist $cont_stats {

	*specify the value of X to specify the excel cells to use 
	local x = `x' + 1

	*put it in the same sheet  
	putexcel set filename, sheet(all) modify

	*use summarise command to get SD, mean and sample number
	sum `pheno'
	
	*store stats under locals and export them to excel file 
	local n_sample = r(N)
	local sd = r(sd)
	local n_mean = r(mean)
	local name : var label `pheno'

	putexcel A`x'="`name'" B`x'=`n_sample' C`x'=`n_mean' D`x'=`sd'

}

***********Binary variables 

*set the cell count to the number of the last cell 
*set up local macro of all binary variables - including maternal smoking 
global bin_stats $bin_pheno sex maternal_smoke 

local x = 22
foreach pheno of varlist $bin_stats {

	*specify the value of X to specify the excel cells to use 
	local x = `x' + 1

	*call the descriptive stats file
	putexcel set filename, sheet(all) modify

	*tab binary outcomes with the matcell option so that the frequency values are stored in a matrix called results
	tab `pheno', matcell(results)
	
	*Pull out relevant results from the matrix created above to store in the excel file 
	local n_sample = r(N)
	local cases = results[2,1]
	local controls = results[1,1]
	local name : var label `pheno'

	putexcel A`x'="`name'" B`x'=`n_sample' E`x'=`controls' G`x'=`cases' 

}




************************* Association with maltreatment ************************

*Set up results file 
*One file, different sheet for each phenotype

foreach pheno in $pheno_risk {

	putexcel set filename, sheet(`pheno') modify
	putexcel A1="Exposure" B1="Outcome" C1="Scale" D1="Beta/OR" ///
		E1="LCI" F1="UCI" G1="P Value" H1="Sample Size" 
		
}

*Create a loop with both linear and logistic regression of maltreatment on phenotypical variable 

*Set macros for the ones being used for linear and logistic regressions
global linear_reg sd_bmi sd_alcohol sd_ldl sd_smoking sd_sbp af diabetes_t2 chd stroke
global ln_reg sd_ln_bmi sd_ln_sbp sd_ln_a_smoking sd_ln_a_alcohol sd_ln_ldl af diabetes_t2 chd stroke
local n : word count $linear_reg

*loop will run an iteration for each of the variables 
forvalues i = 1/`n' {
	
	*specify standardised and log variables, which are on position `i' in the macros, as well as the raw variable - make sure they are all in the same order 
	local linear : word `i' of $linear_reg
	local ln : word `i' of $ln_reg
	local pheno : word `i' of $pheno_risk
	
	*specify the value of X to specify the excel cells to use 
	local x = 2
	
	*specify the file for the results
	putexcel set filename, sheet(`pheno') modify
	
	*linear regression 
	regress `linear' mal_score age sex if touse_`pheno' == 1
	
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta = results[1,1]
	local lci = results[5,1]
	local uci = results[6,1]
	local p_value = results[4,1]
	local out_label : var label `pheno'
	local exp_label : var label mal_score
	local sample_n = e(N)
	
	*Tell excel which values are being stored 
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="Additive" D`x'=`beta' E`x'=`lci' F`x'=`uci' G`x'=`p_value' H`x'=`sample_n'  
	
	local x = `x' + 1

	*logistic regression 
	*add conditions so that the appropriate regression is done according to whether the variable is continuous or binary  - the regress command will be run for the first 5 variables which are continous and the rest of the iterations will use the logistic command as they are binary variables 
	if (`i' <= 5) regress `ln' mal_score age sex if touse_`pheno' == 1 
	else logistic `ln' mal_score age sex if touse_`pheno' == 1 
	
	*store results of whichever regression was run on a matrix to pull the relevants ones
	matrix results2 = r(table)
	local beta = results2[1,1]
	local lci = results2[5,1]
	local uci = results2[6,1]
	local p_value = results2[4,1]
	local out_label : var label `pheno'
	local exp_label : var label mal_score
	local sample_n = e(N)
	
	
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="Multiplicative" D`x'=`beta' E`x'=`lci' F`x'=`uci' G`x'=`p_value' H`x'=`sample_n'
	
}














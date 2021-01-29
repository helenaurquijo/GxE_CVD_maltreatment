* p5e-8 additive - main analyses + sex differences + adjusting for covariate interactions


*Creating maltreatment score based on how many different exposures to maltreatment 

gen mal_score =  physical_abuse + sexual_abuse + emotional_abuse + emotional_neglect + physical_neglect

label var mal_score "Maltreatment score"

*Exclude from sample individuals without maltreatment or GRS data 
mark touse
markout touse mal_score grs_*
keep if touse == 1 

*Create global macros for the GRS, PCs and confounders
*maternal smoke and number of siblings only used for sensitivity analyses - maybe best to set a local since they will differ across the different analyses (main and sensitivity)
global all_grs grs_*
global PCs PC_1-PC_40
global confound sex age PC_1-PC_40


*Standardise GRS 
foreach var of varlist $all_grs {
	
	egen sd_`var' = std(`var')
}

*Create global macros for phenotypic variables
global pheno_risk bmi ldl_c sbp csi drinks_per_week Stroke_inc CHD_inc af_inc t2_inc

*Create global macros for SD GRS at p=5x10-8
global primary_grs sd_grs_af_5e8 sd_grs_alcohol_5e8 sd_grs_bmi_5e8 sd_grs_chd_5e8 sd_grs_diabetes_t2_5e8 sd_grs_ldl_5e8 sd_grs_stroke_5e8

*Label each GRS with phenotype and GRS threshold
lab var sd_grs_alcohol_5e8 "Drinks per week P=5x10-8 (SD)"
lab var sd_grs_af_5e8 "Atrial fibrillation P=5x10-8 (SD)"
lab var sd_grs_bmi_5e8 "BMI P=5x10-8 (SD)"
lab var sd_grs_chd_5e8 "CHD P=5x10-8 (SD)"
lab var sd_grs_diabetes_t2_5e8 "Diabetes (T2) P=5x10-8 (SD)"
lab var sd_grs_ldl_5e8 "LDL-C P=5x10-8 (SD)"
lab var sd_grs_stroke_5e8 "Stroke P=5x10-8 (SD)"


********** Interaction terms 

*Interaction variable for all GRS at p=5e-8 and maltreatment (GRS*maltreatment)
foreach var of varlist $primary_grs {

	gen `var'_int = mal_score * `var'
}

*Interaction term for GRSs and confounders (GRS*confounders) - covariate interactions
foreach var of varlist $primary_grs {

	gen `var'_int_s = sex * `var'
	gen `var'_int_a = age * `var'
	
	foreach pc of varlist $PCs {
	
		gen `var'_int_`pc' = `var' * `pc'
	}
	
}


*Interaction variable for all confounders and maltreatment (Maltreatment*confounders) - covariate interactions
foreach var of varlist $confound {

	gen m_int_`var' = mal_score * `var'
} 

*Create global macro of the interaction terms of confounders with maltreatment (covariate interactions)
global m_int_confound m_int_sex m_int_age m_int_PC_*

*Rename variables
*Match the phenotype to the name of respective GRS
rename drinks_per_week alcohol 
rename csi smoking 
rename t2_inc diabetes_t2
rename CHD_inc chd 
rename af_inc af
rename Stroke_inc stroke 
rename ldl_c ldl

*Re-specify the global macro for the phenotype variables 
global pheno_risk bmi alcohol ldl af smoking sbp chd diabetes_t2  stroke 

*Create touse variable for each of the phenotypes 
foreach var of varlist $pheno_risk {

	mark touse_sd_grs_`var'_5e8
	markout touse_sd_grs_`var'_5e8 mal_score age sex `var'
}


*Specify global macro for continuous phenotypes 
global cont_pheno bmi alcohol ldl smoking sbp

*Standardise continuous variables
foreach var of varlist $cont_pheno {

	egen sd_`var'= std(`var')
}


*********************** Setting up excel results file **************************

*one file for each GRS threshold, different sheet for each phenotype
foreach exp in $primary_grs {

	putexcel set filename, sheet(`exp') modify
	putexcel A1="Exposure" B1="Outcome" C1="Interaction" D1="Scale" ///
		E1="Beta/OR" F1="LCI" G1="UCI" H1="P Value" I1 = "" ///
		J1 = "Sample Size" K1 = "Number of cases" L1 = "Mal Beta" ///
		M1 = "Mal LCI" N1 = "Mal UCI" O1="Mal P value "
		
}

  
******************************************************************************** 
							* Asessing effect of GRSs *
********************************************************************************

*Effect on phenotype per SD increase in GRS

*SBP and smoking samples run separately 

*Specificy local macro for the variables to include in the following loop - all variables except for SMP and smoking 

local pheno_out sd_bmi sd_ldl sd_alcohol chd stroke diabetes_t2 af 
local primary_grs sd_grs_bmi_5e8 sd_grs_ldl_5e8 sd_grs_alcohol_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 sd_grs_diabetes_t2_5e8 sd_grs_af_5e8

*Local macro that counts the number of variables in pheno_out, only calls one of the macros but works on assumption that there is the same number of variables and they are both in the same order - so variables on both macros have to match in order and name 
local n : word count `pheno_out'

*Loop through values of n (number of phenotypes)
forvalues i = 1/`n' {

	*specificy exposure and outcome variables, which are on position `i' in the macros
	local out : word `i' of `pheno_out'
	local exp : word `i' of `primary_grs'
	
	*specify the value of X to specify the excel cells to use 
	local x = 2
	

	*specify which excel sheet this is going into 
	putexcel set filename, sheet(`exp') modify
	
		*****Run the linear regression model 
		*can call the touse variable respective to each of the variables as they have been named just like the phenotype 
	regress `out' `exp' mal_score $confound if touse_`exp' == 1
	 
	
	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	local beta = results[1,1]
	local lci = results[5,1]
	local uci = results[6,1]
	local p_value = results[4,1]
	local out_label : var label `out'
	local exp_label : var label `exp'
	local sample_n = e(N)
	
	*Tell excel which values are being stored 
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="NONE" D`x'="Difference" E`x'=`beta' F`x'=`lci' G`x'=`uci' H`x'=`p_value' J`x'=`sample_n'
	
}

************************ According to sex 

*Set up local macros and word count as before 
local pheno_out sd_bmi sd_ldl sd_alcohol chd stroke diabetes_t2 af 
local primary_grs sd_grs_bmi_5e8 sd_grs_ldl_5e8 sd_grs_alcohol_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 sd_grs_diabetes_t2_5e8 sd_grs_af_5e8
local n : word count `pheno_out'

*Call the 'levels' of sex i.e. female and male, and create a local macro sith 
levelsof sex, local(levels)

forvalues i = 1/`n' {

*specificy exposure and outcome variables, which are on position `i' in the macros
local out : word `i' of `pheno_out'
local exp : word `i' of `primary_grs'

*specify the value of X to specify the excel cells to use 
local x = 5

foreach l of local levels {

	*specify that for each iteration of the loop we are adding 1 to the value of x, so each result goes in the cell below the previous one 
	local x = `x' + 1
	
	*specify the excel file 
	putexcel set filename, sheet(`exp') modify
	
	regress `out' `exp' mal_score $confound if sex == `l' & touse_`exp' == 1

	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	local beta = results[1,1]
	local lci = results[5,1]
	local uci = results[6,1]
	local p_value = results[4,1]
	local out_label : var label `out'
	local exp_label : var label `exp'
	local sample_n = e(N)
	
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="`l' sex " D`x'="Difference" E`x'=`beta' F`x'=`lci' G`x'=`uci' H`x'=`p_value' J`x'=`sample_n'
	
}

}


********************************************************************************
					* Testing for additive interactions *
********************************************************************************

*same local macros pheno_out, primary_grs and n apply
*need to call them again 
local pheno_out sd_bmi sd_ldl sd_alcohol chd stroke diabetes_t2 af 
local primary_grs sd_grs_bmi_5e8 sd_grs_ldl_5e8 sd_grs_alcohol_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 sd_grs_diabetes_t2_5e8 sd_grs_af_5e8
local n : word count `pheno_out'

*create similar loop but this time looking at interactions 
forvalues i = 1/`n' {

	local out : word `i' of `pheno_out'
	local exp : word `i' of `primary_grs'

	*specify the value of X to specify the excel cells to use, need to one row below so that it doesn't overlap
	local x = 3

	*specify which excel sheet this is going into - same one as before 
	putexcel set filename, sheet(`exp') modify
	
	*****Run the linear regression model but this time with the interaction term 
	*can call the interaction term, as well as interactions of maltreatment with confounders (m_int) and GRS with confounders (a_int, s_int, )
	regress `out' `exp' mal_score `exp'_int $confound $m_int_confound `exp'_int_a `exp'_int_s `exp'_int_PC* if touse_`exp' == 1
	
	*Store the results in a matrix and call local macros to export results to excel
	matrix results = r(table)
	local beta_int = results[1,3]
	local lci_int = results[5,3]
	local uci_int = results[6,3]
	local p_value = results[4,3]
	local mal_beta = results[1,2]
	local lci_mal = results[5,2]
	local uci_mal = results[6,2]
	local mal_p = results[4,2]
	local out_label : var label `out'
	local exp_label : var label `exp'
 
	
	*Tell excel which values are being stored 		
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="Interaction+conf" D`x'="Additive" E`x'=`beta_int' F`x'=`lci_int' G`x'=`uci_int' H`x'=`p_value' L`x'=`mal_beta' M`x'=`lci_mal' N`x'=`uci_mal' O`x'=`mal_p'

	*Now going to run the same regression but without the extra interaction terms to see what it does to the results, store the same results 
	local x = `x' + 1
	
	regress `out' `exp' mal_score `exp'_int $confound if touse_`exp' == 1
	matrix results = r(table)
	local beta_int = results[1,3]
	local lci_int = results[5,3]
	local uci_int = results[6,3]
	local p_value = results[4,3]
	local mal_beta = results[1,2]
	local lci_mal = results[5,2]
	local uci_mal = results[6,2]
	local mal_p = results[4,2]
	local out_label : var label `out'
	local exp_label : var label `exp'
	
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="Interaction" D`x'="Additive" E`x'=`beta_int' F`x'=`lci_int' G`x'=`uci_int' H`x'=`p_value' L`x'=`mal_beta' M`x'=`lci_mal' N`x'=`uci_mal' O`x'=`mal_p'
	
}

************************ According to sex 

*Set up local macros and word count as before 
local pheno_out sd_bmi sd_ldl sd_alcohol chd stroke diabetes_t2 af 
local primary_grs sd_grs_bmi_5e8 sd_grs_ldl_5e8 sd_grs_alcohol_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 sd_grs_diabetes_t2_5e8 sd_grs_af_5e8
local n : word count `pheno_out'

*Call the 'levels' of sex i.e. female and male, and create a local macro sith 
levelsof sex, local(levels)

forvalues i = 1/`n' {

*specificy exposure and outcome variables, which are on position `i' in the macros
local out : word `i' of `pheno_out'
local exp : word `i' of `primary_grs'

*specify the value of X to specify the excel cells to use 
local x = 7

foreach l of local levels {

	*specify that for each iteration of the loop we are adding 1 to the value of x, so each result goes in the cell below the previous one 
	local x = `x' + 1
	
	*specify the excel file 
	putexcel set filename, sheet(`exp') modify
	
	regress `out' `exp' mal_score `exp'_int $confound if sex == `l' & touse_`exp' == 1

	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta_int = results[1,3]
	local lci_int = results[5,3]
	local uci_int = results[6,3]
	local p_value = results[4,3]
	local mal_beta = results[1,2]
	local lci_mal = results[5,2]
	local uci_mal = results[6,2]
	local mal_p = results[4,2]
	local out_label : var label `out'
	local exp_label : var label `exp'
	
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="`l' sex " D`x'="Additive" E`x'=`beta_int' F`x'=`lci_int' G`x'=`uci_int' H`x'=`p_value' L`x'=`mal_beta' M`x'=`lci_mal' N`x'=`uci_mal' O`x'=`mal_p'
	
}

}

/*Assessing difference between sexes

*same local macros pheno_out, primary_grs and n apply
*need to call them again 
local pheno_out sd_bmi sd_ldl sd_alcohol chd stroke diabetes_t2 af 
local primary_grs sd_grs_bmi_5e8 sd_grs_ldl_5e8 sd_grs_alcohol_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 sd_grs_diabetes_t2_5e8 sd_grs_af_5e8
local n : word count `pheno_out'

*create similar loop but this time looking at interactions 
forvalues i = 1/`n' {

	local out : word `i' of `pheno_out'
	local exp : word `i' of `primary_grs'

	*specify the value of X to specify the excel cells to use, need to one row below so that it doesn't overlap
	local x = 10

	*specify which excel sheet this is going into - same one as before 
	putexcel set filename, sheet(`exp') modify

	*Run the regression twice - once with the the previously tested interaction and then with a three way interaction with sex 
	regress `out' c.`exp'##c.mal_score $confound if touse_`exp' == 1 
	est store A
	regress `out' c.`exp'##c.mal_score##i.sex $confound if touse_`exp' == 1 
	est store B 
	
	*Likelihood-ratio test to get p value for the effect of sex and export it 
	lrtest A B
	local p_val = r(p)
	
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'= "Sex difference" D`x'= "Additive" H`x'= `p_val'
	
} */

*REVISED METHOD FOR THE INTERACTION BY SEX COEFFICIENTS AND P VALUES - the other one was picking up all possible interactions rather than just the 3 way one 

*same local macros pheno_out, primary_grs and n apply
*need to call them again 
local pheno_out sd_bmi sd_ldl sd_alcohol chd stroke diabetes_t2 af 
local primary_grs sd_grs_bmi_5e8 sd_grs_ldl_5e8 sd_grs_alcohol_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 sd_grs_diabetes_t2_5e8 sd_grs_af_5e8
local n : word count `pheno_out'


forvalues i = 1/`n' {

	local out : word `i' of `pheno_out'
	local exp : word `i' of `primary_grs'

	local x = 11
	
	*specify which excel sheet this is going into - same one as before 
	putexcel set filename, sheet(`exp') modify

	*Run the regression twice - once with the the previously tested interaction and then with a three way interaction with sex 
	
	regress `out' c.`exp'##c.mal_score##i.sex $confound if touse_`exp' == 1 
	
	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta_3 = results[1,11]
	local lci_3 = results[5,11]
	local uci_3 = results[6,11]
	local p_val_3 = results[4,11]

	
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'="Sex difference" D`x'="Additive" E`x'=`beta_3' F`x'=`lci_3' G`x'=`uci_3' H`x'=`p_val_3' 
	
}

**************** Assessing variance explained by GRS


*Set up a file - putting it all in the same sheet
putexcel set filename, sheet(all) modify
putexcel A1="Exposure" B1="Outcome" C1="R squared" D1="Adjusted R squared" 


*Create local macros for the variables being used
global pheno_out sd_bmi sd_alcohol sd_ldl af diabetes_t2 chd stroke  
global primary_grs sd_grs_bmi_5e8 sd_grs_alcohol_5e8 sd_grs_ldl_5e8 sd_grs_af_5e8 sd_grs_diabetes_t2_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8 
local n : word count $pheno_out

*set this local macro so that the loop first writes results onto the second row of the excel file
local x = 2
forvalues i = 1/`n' {

	*specify the value of X to specify the excel cells to use - where to start
	

	*specify exposure and outcome variables, which are on position `i' in the macros
	local outcome : word `i' of $pheno_out
	local exp : word `i' of $primary_grs
	

	*specify which excel sheet this is going into 
	putexcel set filename, sheet(all) modify
	
		*****Run the linear regression model 
		*can call the touse variable respective to each of the variables as they have been named just like the phenotype 
	regress `outcome' `exp' $confound if touse_`exp' == 1
	
	*Pull out relevant results from matrix to store in the excel file 
	local squared = e(r2)
	local adjsquared = e(r2_a)
	local out_label : var label `outcome'
	local exp_label : var label `exp'
	
	*Tell excel which values are being stored 
	putexcel A`x'="`exp_label'" B`x'="`out_label'" C`x'=`squared' D`x'=`adjsquared'
	
	*add one so that the next loop writes on the row below
	local x = `x' + 1
}



******************************************************************************** 
						  * Split Sample Analysis *
********************************************************************************

*Smoking and systolic blood pressure need to be done separately as the GWAS to create the GRS was done with data from the Biobank (Split sample)

********************************** Smoking ************************************* 

************** Effect of GRS

*This is on a loop but doesn't need to be 

foreach out in sd_smoking {

local x = 1
local x = `x' + 1

	*Set up a separate excel file for the smoking results 
	putexcel set filename, sheet(all) modify
	putexcel A1="sample" B1="out" C1="log_or" D1="lci" E1="uci" F1="n" G1="mal_smoking" H1="mal_beta" I1="mal_lci" J1="mal_uci" K1 = "sex"
	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store
	regress `out' sd_grs_smoking_5e8_gwas2 mal_score $confound if sample == 1 & touse_sd_grs_smoking_5e8 == 1
	
	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas2]-1.96*_se[sd_grs_smoking_5e8_gwas2]
	local uci = _b[sd_grs_smoking_5e8_gwas2]+1.96*_se[sd_grs_smoking_5e8_gwas2]
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas2] D`x'=`lci' E`x'=`uci' G`x'="NONE"
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_smoking_5e8_gwas1 mal_score $confound if sample == 2 & touse_sd_grs_smoking_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas1]-1.96*_se[sd_grs_smoking_5e8_gwas1]
	local uci = _b[sd_grs_smoking_5e8_gwas1]+1.96*_se[sd_grs_smoking_5e8_gwas1]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas1] D`x'=`lci' E`x'=`uci' G`x'="NONE"
	
}

********** Effect of GRS according to sex 

levelsof sex, local(levels)

local x = 6

foreach out in sd_smoking {


foreach l of local levels {

	putexcel set filename, sheet(all) modify
 	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store
	regress `out' sd_grs_smoking_5e8_gwas2 mal_score $confound if sex == `l' & sample == 1 & touse_sd_grs_smoking_5e8 == 1
	
	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas2]-1.96*_se[sd_grs_smoking_5e8_gwas2]
	local uci = _b[sd_grs_smoking_5e8_gwas2]+1.96*_se[sd_grs_smoking_5e8_gwas2]
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas2] D`x'=`lci' E`x'=`uci' G`x'="NONE" K`x'="`l' sex"
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_smoking_5e8_gwas1 mal_score $confound if sex == `l' & sample == 2 & touse_sd_grs_smoking_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas1]-1.96*_se[sd_grs_smoking_5e8_gwas1]
	local uci = _b[sd_grs_smoking_5e8_gwas1]+1.96*_se[sd_grs_smoking_5e8_gwas1]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas1] D`x'=`lci' E`x'=`uci' G`x'="NONE" K`x'="`l' sex"
	
	local x = `x' + 1
	
}

}

****************** Interaction of smoking GRS with maltreatment***************** 

*Create interaction term of GRS with maltreatment 
gen sd_grs_smoking_5e8_gwas2_int = sd_grs_smoking_5e8_gwas2 * mal_score if sample==1
gen sd_grs_smoking_5e8_gwas1_int = sd_grs_smoking_5e8_gwas1 * mal_score if sample==2

*Create interaction terms for GRS with confounders 
/*foreach var of varlist $confound {

	gen grs_int_2_`var' = `var' * sd_grs_smoking_5e8_gwas2 if sample == 1
	gen grs_int_1_`var' = `var' * sd_grs_smoking_5e8_gwas1 if sample == 2

}

*create global macro for all the confounder interactions with smoking GRS 
global smoking_grs_int_2 grs_int_2_age grs_int_2_sex grs_int_2_PC*
global smoking_grs_int_1 grs_int_1_age grs_int_1_sex grs_int_1_PC* */


foreach out in sd_smoking {

local x = 4


	*Call the  excel file for the smoking results 
	putexcel set filename, sheet(all) modify
	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store but including the term for interaction and the terms for interactions with confounders  
	regress `out' sd_grs_smoking_5e8_gwas2 mal_score sd_grs_smoking_5e8_gwas2_int $confound  if sample == 1 & touse_sd_grs_smoking_5e8 == 1

	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out' // not sure if including this one 
	local lci = _b[sd_grs_smoking_5e8_gwas2_int]-1.96*_se[sd_grs_smoking_5e8_gwas2_int]
	local uci = _b[sd_grs_smoking_5e8_gwas2_int]+1.96*_se[sd_grs_smoking_5e8_gwas2_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas2_int] D`x'=`lci' E`x'=`uci'  G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci'
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_smoking_5e8_gwas1 mal_score sd_grs_smoking_5e8_gwas1_int $confound  if sample == 2 & touse_sd_grs_smoking_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas1_int]-1.96*_se[sd_grs_smoking_5e8_gwas1_int]
	local uci = _b[sd_grs_smoking_5e8_gwas1_int]+1.96*_se[sd_grs_smoking_5e8_gwas1_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas1_int] D`x'=`lci' E`x'=`uci'  G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci'
	
}

********** Interaction according to sex 

levelsof sex, local(levels)

local x = 10

foreach out in sd_smoking {


foreach l of local levels {

	putexcel set filename, sheet(all) modify
	putexcel A1="sample" B1="out" C1="log_or" D1="lci" E1="uci" F1="n" G1="mal_smoking" H1="mal_beta" I1="mal_lci" J1="mal_uci"
	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store
	regress `out' sd_grs_smoking_5e8_gwas2 mal_score sd_grs_smoking_5e8_gwas2_int $confound if sex == `l' & sample == 1 & touse_sd_grs_smoking_5e8 == 1
	
	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way
	
	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas2_int]-1.96*_se[sd_grs_smoking_5e8_gwas2_int]
	local uci = _b[sd_grs_smoking_5e8_gwas2_int]+1.96*_se[sd_grs_smoking_5e8_gwas2_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas2_int] D`x'=`lci' E`x'=`uci' G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci' K`x'="`l' sex"
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_smoking_5e8_gwas1 mal_score sd_grs_smoking_5e8_gwas1_int $confound if sex == `l' & sample == 2 & touse_sd_grs_smoking_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_smoking_5e8_gwas1_int]-1.96*_se[sd_grs_smoking_5e8_gwas1_int]
	local uci = _b[sd_grs_smoking_5e8_gwas1_int]+1.96*_se[sd_grs_smoking_5e8_gwas1_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_smoking_5e8_gwas1_int] D`x'=`lci' E`x'=`uci' G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci' K`x'="`l' sex"
	
	local x = `x' + 1
	
}

}

*Assessing the difference between sexes - using method as for other outcomes butonce for every sample 

foreach out in sd_smoking {

	
	local x = 15
	
	*specify which excel sheet this is going into - same one as before 
	putexcel set filename, sheet(all) modify
	putexcel L1 = "beta_3" M1 = "lci_3" N1 = "uci_3"

	*Run the regression twice - once with the the previously tested interaction and then with a three way interaction with sex 
	
	regress `out' c.sd_grs_smoking_5e8_gwas2##c.mal_score##i.sex $confound if sample == 1 & touse_sd_grs_smoking_5e8 == 1
	
	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta_3 = results[1,11]
	local lci_3 = results[5,11]
	local uci_3 = results[6,11]
	
	putexcel A`x'="1" B`x'="`out_label'" L`x'=`beta_3' M`x'=`lci_3'  N`x'=`uci_3'
	
	local x = `x'+ 1
	
	regress `out' c.sd_grs_smoking_5e8_gwas1##c.mal_score##i.sex $confound if sample == 2 & touse_sd_grs_smoking_5e8 == 1
	
	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta_3 = results[1,11]
	local lci_3 = results[5,11]
	local uci_3 = results[6,11]
	
	putexcel A`x'="2" B`x'="`out_label'" L`x'=`beta_3' M`x'=`lci_3'  N`x'=`uci_3'

}

*****Assessing variance explained by PRS - R2
local x = 9

foreach out in sd_smoking sd_sbp {

	local exp = substr("`out'",4,.)

	*specify which excel sheet this is going into 
	putexcel set filename, sheet(all) modify
	
		
	regress `out' sd_grs_`exp'_5e8_gwas2 $confound if touse_sd_grs_`exp'_5e8 == 1 & sample == 1
	
	*Pull out relevant results from memory to store in the excel file 
	local adjsquared = e(r2_a)
	local out_label : var label `out'
	
	*Tell excel which values are being stored 
	putexcel A`x'="5e-8" B`x'="`out_label'" C`x'="1" D`x'=`adjsquared'
	
	*add one so that the next loop writes on the row below
	local x = `x' + 1
	
	regress `out' sd_grs_`exp'_5e8_gwas1 $confound if touse_sd_grs_`exp'_5e8 == 1 & sample == 2
	
	*Pull out relevant results from memory to store in the excel file 
	local adjsquared = e(r2_a)
	local out_label : var label `out'
	
	*Tell excel which values are being stored 
	putexcel A`x'="5e-8" B`x'="`out_label'" C`x'="2" D`x'=`adjsquared'
	
	local x = `x' + 1
}


*************************** Systolic blood pressure ****************************

************** effect of GRS 


foreach out in sd_sbp {

local x = 1
local x = `x' + 1

	*Set up a separate excel file for the sbp results 
	putexcel set filename, sheet(all) modify
	putexcel A1="sample" B1="out" C1="log_or" D1="lci" E1="uci" F1="n" G1="mal_sbp" H1="mal_beta" I1="mal_lci" J1="mal_uci" K1 = "sex"
	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store
	regress `out' sd_grs_sbp_5e8_gwas2 mal_score $confound if sample == 1 & touse_sd_grs_sbp_5e8 == 1
	*again, use maltreatment here?
	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas2]-1.96*_se[sd_grs_sbp_5e8_gwas2]
	local uci = _b[sd_grs_sbp_5e8_gwas2]+1.96*_se[sd_grs_sbp_5e8_gwas2]
	
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas2] D`x'=`lci' E`x'=`uci' G`x'="NONE" 
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_sbp_5e8_gwas1 mal_score $confound if sample == 2 & touse_sd_grs_sbp_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas1]-1.96*_se[sd_grs_sbp_5e8_gwas1]
	local uci = _b[sd_grs_sbp_5e8_gwas1]+1.96*_se[sd_grs_sbp_5e8_gwas1]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas1] D`x'=`lci' E`x'=`uci'  G`x'="NONE" 
	
}

********** Effect of GRS according to sex 

levelsof sex, local(levels)

local x = 6

foreach out in sd_sbp {


foreach l of local levels {

	putexcel set filename, sheet(all) modify
 	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store
	regress `out' sd_grs_sbp_5e8_gwas2 mal_score $confound if sex == `l' & sample == 1 & touse_sd_grs_sbp_5e8 == 1
	
	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas2]-1.96*_se[sd_grs_sbp_5e8_gwas2]
	local uci = _b[sd_grs_sbp_5e8_gwas2]+1.96*_se[sd_grs_sbp_5e8_gwas2]
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas2] D`x'=`lci' E`x'=`uci' G`x'="NONE" K`x'="`l' sex"
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_sbp_5e8_gwas1 mal_score $confound if sex == `l' & sample == 2 & touse_sd_grs_sbp_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas1]-1.96*_se[sd_grs_sbp_5e8_gwas1]
	local uci = _b[sd_grs_sbp_5e8_gwas1]+1.96*_se[sd_grs_sbp_5e8_gwas1]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas1] D`x'=`lci' E`x'=`uci' G`x'="NONE" K`x'="`l' sex"
	
	local x = `x' + 1
	
}

}

**** Interaction of systolic blood pressure GRS with maltretament 

*Create interaction term with maltreatment 
gen sd_grs_sbp_5e8_gwas2_int = sd_grs_sbp_5e8_gwas2 * mal_score if sample==1
gen sd_grs_sbp_5e8_gwas1_int = sd_grs_sbp_5e8_gwas1 * mal_score if sample==2

*Create interaction terms for GRS with confounders 
foreach var of varlist $confound {

	gen sbp_grs_int_2_`var' = `var' * sd_grs_sbp_5e8_gwas2 if sample == 1
	gen sbp_grs_int_1_`var' = `var' * sd_grs_sbp_5e8_gwas1 if sample == 2

}

*create global macro for all the confounder inetractions with SBP GRS 
global sbp_grs_int_2 sbp_grs_int_2_age sbp_grs_int_2_sex sbp_grs_int_2_PC*
global sbp_grs_int_1 sbp_grs_int_1_age sbp_grs_int_1_sex sbp_grs_int_1_PC*


foreach out in sd_sbp {

local x = 4


	*Call the  excel file for the SBP results 
	putexcel set filename, sheet(all) modify
	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store but including the term for interaction and the terms for interactions with confounders  
	regress `out' sd_grs_sbp_5e8_gwas2 mal_score sd_grs_sbp_5e8_gwas2_int $confound  if sample == 1 & touse_sd_grs_sbp_5e8 == 1

	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out' // not sure if including this one 
	local lci = _b[sd_grs_sbp_5e8_gwas2_int]-1.96*_se[sd_grs_sbp_5e8_gwas2_int]
	local uci = _b[sd_grs_sbp_5e8_gwas2_int]+1.96*_se[sd_grs_sbp_5e8_gwas2_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas2_int] D`x'=`lci' E`x'=`uci' G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci'
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_sbp_5e8_gwas1 mal_score sd_grs_sbp_5e8_gwas1_int $confound  if sample == 2 & touse_sd_grs_sbp_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas1_int]-1.96*_se[sd_grs_sbp_5e8_gwas1_int]
	local uci = _b[sd_grs_sbp_5e8_gwas1_int]+1.96*_se[sd_grs_sbp_5e8_gwas1_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas1_int] D`x'=`lci' E`x'=`uci'  G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci'
	
}

********** Interaction according to sex 

levelsof sex, local(levels)

local x = 10

foreach out in sd_sbp {


foreach l of local levels {

	putexcel set filename, sheet(all) modify
	putexcel A1="sample" B1="out" C1="log_or" D1="lci" E1="uci" F1="n" G1="mal_sbp" H1="mal_beta" I1="mal_lci" J1="mal_uci"
	
	*Run the linear regression on sample 1 using the GRS obstained from sample 2 and store
	regress `out' sd_grs_sbp_5e8_gwas2 mal_score sd_grs_sbp_5e8_gwas2_int $confound if sex == `l' & sample == 1 & touse_sd_grs_sbp_5e8 == 1
	
	
	*Results have been stored - interested in effect estimates and confidences intervals so won't store in matrix
	*regression coefficients are stored under _b and can accessed that way 
	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas2_int]-1.96*_se[sd_grs_sbp_5e8_gwas2_int]
	local uci = _b[sd_grs_sbp_5e8_gwas2_int]+1.96*_se[sd_grs_sbp_5e8_gwas2_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="1" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas2_int] D`x'=`lci' E`x'=`uci' G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci' K`x'="`l' sex"
	
	*Run regression on sample 2 using the GRS obtained from sample 1 
	regress `out' sd_grs_sbp_5e8_gwas1 mal_score sd_grs_sbp_5e8_gwas1_int $confound if sex == `l' & sample == 2 & touse_sd_grs_sbp_5e8 == 1
	
	* Add 1 to the value of x to write on the cell row below 
	local x = `x' + 1

	local out_label : var label `out'
	local lci = _b[sd_grs_sbp_5e8_gwas1_int]-1.96*_se[sd_grs_sbp_5e8_gwas1_int]
	local uci = _b[sd_grs_sbp_5e8_gwas1_int]+1.96*_se[sd_grs_sbp_5e8_gwas1_int]
	local mal_lci = _b[mal_score]-1.96*_se[mal_score]
	local mal_uci = _b[mal_score]+1.96*_se[mal_score]
	
	putexcel A`x'="2" B`x'="`out_label'" C`x'=_b[sd_grs_sbp_5e8_gwas1_int] D`x'=`lci' E`x'=`uci' G`x'="Interaction" H`x'=_b[mal_score] I`x'=`mal_lci' J`x'=`mal_uci' K`x'="`l' sex"
	
	local x = `x' + 1
	
}

}

*Assessing the difference between sexes - using method as for other outcomes but once for every sample 

foreach out in sd_sbp {

	
	local x = 15
	
	*specify which excel sheet this is going into - same one as before 
	putexcel set filename, sheet(all) modify
	putexcel L1 = "beta_3" M1 = "lci_3" N1 = "uci_3"

	*Run the regression twice - once with the the previously tested interaction and then with a three way interaction with sex 
	
	regress `out' c.sd_grs_sbp_5e8_gwas2##c.mal_score##i.sex $confound if sample == 1 & touse_sd_grs_sbp_5e8 == 1
	
	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta_3 = results[1,11]
	local lci_3 = results[5,11]
	local uci_3 = results[6,11]
	
	putexcel A`x'="1" B`x'="`out_label'" L`x'=`beta_3' M`x'=`lci_3'  N`x'=`uci_3'
	
	local x = `x'+ 1
	
	regress `out' c.sd_grs_sbp_5e8_gwas1##c.mal_score##i.sex $confound if sample == 2 & touse_sd_grs_sbp_5e8 == 1
	
	*Store the results in a matrix 
	matrix results = r(table)
	
	*Pull out relevant results from matrix to store in the excel file 
	matrix results = r(table)
	local beta_3 = results[1,11]
	local lci_3 = results[5,11]
	local uci_3 = results[6,11]
	
	putexcel A`x'="2" B`x'="`out_label'" L`x'=`beta_3' M`x'=`lci_3'  N`x'=`uci_3'

}

****************** Tab sample size 

********Tab sample size for smking and SBP analyses - count of sample 1 and 2 together 
*This stores the sample size in main results file rather than the separate ones

foreach touse in touse_sd_grs_smoking_5e8 touse_sd_grs_sbp_5e8 {

	local sheet = substr("`touse'",7,.)
	
	putexcel set filename, sheet(`sheet') modify 
	
	tab mal_score if `touse' == 1, matcell(numbers)
	
	local n_0 = numbers[1,1]
	local n_1 = numbers[2,1]
	local n_2 = numbers[3,1]
	local n_3 = numbers[4,1]
	local n_4 = numbers[5,1]
	local n_5 = numbers[6,1]
	local n_all = sum(`n_0'+`n_1'+`n_2'+`n_3'+`n_4'+`n_5')
		
	
	putexcel J2=`n_all' J3=`n_0' J4=`n_1' J5=`n_2' J6=`n_3' J7=`n_4' J8=`n_5'

}

**************Tab number of cases
***Tab number of cases for all binary outcomes 

local binary_out af diabetes_t2 chd stroke 
local binary_primary_grs sd_grs_af_5e8 sd_grs_diabetes_t2_5e8 sd_grs_chd_5e8 sd_grs_stroke_5e8

local n : word count `binary_out'

forvalues i = 1/`n' {

	local outcome : word `i' of `binary_out'
	local exp : word `i' of `binary_primary_grs'

	putexcel set filename, sheet(`exp') modify 

	*Tabing number like this (per score) though we will probably present this data as cases in people that were exposed maltreatment as those who were regardless of score (as we are treating maltreatment as continuous)
		tab `outcome' mal_score if touse_`exp' == 1, matcell(numbers)
	
		local n_0 = numbers[2,1]
		local n_1 = numbers[2,2]
		local n_2 = numbers[2,3]
		local n_3 = numbers[2,4]
		local n_4 = numbers[2,5]
		local n_5 = numbers[2,6]
		local n_all = `n_1'+`n_2'+`n_3'+`n_4'+`n_5'
	
	
		putexcel K2=`n_all' K3=`n_0' K4=`n_1' K5=`n_2' K6=`n_3' K7=`n_4' K8=`n_5'

}


******************************************************************************** 
						* Meta-analysing split samples *
********************************************************************************

*********************** 2 sets of results from SMOKING 

*Import the sets of results from excel 
import excel "filename", sheet("all") firstrow clear

*Sample of the variable was defined when exporting the results but as a string, now convert the strings into numbers
*call the variable numbers and replace the variable with numeric version
destring sample, replace
lab var sample "Sample"
lab def sample 1 "Sample 1" 2 "Sample 2"
lab val sample sample 

lab var n "N"


*Create a varaible to identify the sex for the sex-stratified coefficients 
generate sex_level=.
replace sex_level = 1 if sex=="0 sex"
replace sex_level = 2 if sex=="1 sex"

*Create variable to identify the 3-way values 
gen sex_int = 1 if beta_3 !=.

*Create a string variable to know which ones are the interaction values - included in the analyses loops 
*Create a variable from the former string variable to distinguish the non-interactive effects and the interactive ones so that they are correctly specified in the meta analyses
gen mal_smoking_interact = 0 if mal_smoking=="NONE"
replace mal_smoking_interact = 1 if mal_smoking=="Interaction"

	*****Run the meta analysis
	*First, non-interactive coefficients 
	metan log_or lci uci if mal_smoking_interact==0 & sex_level==., nograph
	*Specify the excel file to store the results and the names of the columns
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	putexcel A1="Exposure" B1="Outcome" C1="Interaction" D1="Scale" ///
		E1="Beta/OR" F1="LCI" G1="UCI" H1="P Value" I1 = "P value for interaction (with maltreatment)" J1 = "Sample size" 
	*Specify x for the cell row number 
	local x = 2
	*Create local macros for the results to store
	local beta = (r(ES))
	local lci = (r(ci_low))
	local uci = (r(ci_upp))
	local p_val = r(p_z)
	*Store the results
	putexcel A`x'="Smoking combined P=5x10-8 (SD)" B`x'="Smoking" C`x'="NONE" D`x'="Difference" E`x'=`beta' F`x'=`lci' G`x'=`uci' H`x'=`p_val'
	
	*Next, the sex stratified non-interactive coefficients 
	levelsof sex_level, local(levels)
	local x = 5

foreach l of local levels {
	
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	
	metan log_or lci uci if sex_level==`l' & mal_smoking_interact==0, lcols(sample) effect(OR) null(1) xlabel(-0.0001, -0.001, -0.01, 0.0) astext(50) dp(6) nograph
	
	local x = `x'+1
	local beta_s = (r(ES))
	local lci_s = (r(ci_low))
	local uci_s = (r(ci_upp))
	local p_val_s = r(p_z)	
	putexcel A`x'="Smoking combined P=5x10-8 (SD)" B`x'="Smoking" C`x'="`l' sex" D`x'="Difference" E`x'=`beta_s' F`x'=`lci_s' G`x'=`uci_s' H`x'=`p_val_s'
	
	
}


	*And now the interactive coefficients
	*This time we specify that we are interested in the interactions hence mal_smoking_interact==1
	metan log_or lci uci if mal_smoking_interact==1 & sex_level ==. , nograph
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	
	*Specify the local for excel row number and to store result values
	local x=3
	local beta_int = r(ES)
	local lci_int = r(ci_low)
	local uci_int = r(ci_upp)
	local p_value = r(p_z)

	*Store the results in the excel file
	*Moved the p value to column I so it matches the headings defined above
	putexcel A`x'="Smoking combined P=5x10-8 (SD)" B`x'="Smoking" C`x'="Interaction" D`x'="Additive" E`x'=`beta_int' F`x'=`lci_int' G`x'=`uci_int' I`x'=`p_value'
	
	*Then the interactive coefficient but sex stratified 
	levelsof sex_level, local(levels)
	local x = 7
	
	foreach l of local levels {
	
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	
	metan log_or lci uci if sex_level==`l' & mal_smoking_interact==1, lcols(sample) effect(OR) null(1) xlabel(-0.0001, -0.001, -0.01, 0.0) astext(50) dp(6) nograph
	
	local x = `x'+1
	local beta_s = (r(ES))
	local lci_s = (r(ci_low))
	local uci_s = (r(ci_upp))
	local p_val_s = r(p_z)	
	putexcel A`x'="Smoking combined P=5x10-8 (SD)" B`x'="Smoking" C`x'="`l' sex" D`x'="Additive" E`x'=`beta_s' F`x'=`lci_s' G`x'=`uci_s' I`x'=`p_val_s'
	
	
}
	
	*Now want to meta-analyse the maltreatment coeefficient too - only interaction results
	metan mal_beta mal_lci mal_uci if mal_smoking_interact==1 & sex_level==., nograph
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	
	*Specify the local for excel row number and to store result values
	local x = 4
	local mal_beta = r(ES)
	local mal_lci = r(ci_low)
	local mal_uci = r(ci_upp)
	local mal_p = r(p_z)
	
	*Store the results in the excel file
	putexcel A`x'="Maltreatment score" B`x'="Smoking" C`x'="Interaction" D`x'="Additive" E`x'=`mal_beta' F`x'=`mal_lci' G`x'=`mal_uci' I`x'=`mal_p'
	
	*Then the maltreatment coefficient but sex stratified 
	levelsof sex_level, local(levels)
	local x = 9
	
	foreach l of local levels {
	
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	
	metan mal_beta mal_lci mal_uci if sex_level==`l', lcols(sample) effect(OR) null(1) xlabel(-0.0001, -0.001, -0.01, 0.0) astext(50) dp(6) nograph
	
	local x = `x'+1
	local beta_s = (r(ES))
	local lci_s = (r(ci_low))
	local uci_s = (r(ci_upp))
	local p_val_s = r(p_z)	
	putexcel A`x'="Maltreatment score" B`x'="Smoking" C`x'="`l' sex" D`x'="Additive" E`x'=`beta_s' F`x'=`lci_s' G`x'=`uci_s' H`x'=`p_val_s'
	
	
}

	*And then the 3-way interaction 
	metan beta_3 lci_3 uci_3 if sex_int == 1, nograph 
	putexcel set filename, sheet(sd_grs_smoking_5e8) modify
	
	*Specify the local for excel row number and to store result values
	local x = 12
	local b_3 = r(ES)
	local l_3 = r(ci_low)
	local u_3 = r(ci_upp)
	local p_3 = r(p_z)
	
	*Store the results in the excel file
	putexcel  C`x'=" 3 way Interaction" D`x'="Additive" E`x'=`b_3' F`x'=`b_3' G`x'=`b_3' I`x'=`p_3'
	


*Now the same analysis but for the systolic blood pressure results 

************************** 2 sets of results from SBP 

*Import the sets of results from excel 
import excel "filename", sheet("all") firstrow clear

*Sample of the variable was defined when exporting the results but as a string, now convert the strings into numbers
*call the variable numbers and replace the variable with numeric version
destring sample, replace
lab var sample "Sample"
lab def sample 1 "Sample 1" 2 "Sample 2"
lab val sample sample // not sure what this line is doing 

lab var n "N"

*Create a varaible to identify the sex for the sex-stratified coefficients 
generate sex_level=.
replace sex_level = 1 if sex=="0 sex"
replace sex_level = 2 if sex=="1 sex"


*Create a variable to identify the 3 way interaction results 
gen sex_int = 1 if beta_3 !=.

*Create a string variable to know which ones are the interaction values - included in the analyses loops 
*Create a variable from the former string variable to distinguish the non-interactive effects and the interactive ones so that they are correctly specified in the meta analyses
gen mal_sbp_interact = 0 if mal_sbp=="NONE"
replace mal_sbp_interact = 1 if mal_sbp=="Interaction"

	*****Run the meta analysis
	*First, non-interactive coefficients 
	metan log_or lci uci if mal_sbp_interact==0 & sex_level==., nograph
	*Specify the excel file to store the results and the names of the columns
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify
	putexcel A1="Exposure" B1="Outcome" C1="Interaction" D1="Scale" ///
		E1="Beta/OR" F1="LCI" G1="UCI" H1="P Value" I1 = "P value for interaction (with maltreatment)" J1 = "Sample size" 
	*Specify x for the cell row number 
	local x = 2
	*Create local macros for the results to store
	local beta = (r(ES))
	local lci = (r(ci_low))
	local uci = (r(ci_upp))
	local p_val = r(p_z)
	*Store the results
	putexcel A`x'="SBP combined P=5x10-8 (SD)" B`x'="SBP" C`x'="NONE" D`x'="Difference" E`x'=`beta' F`x'=`lci' G`x'=`uci' H`x'=`p_val'

	*Next, the sex stratified coefficients 
	levelsof sex_level, local(levels)
	local x = 5

foreach l of local levels {
	
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify
	
	metan log_or lci uci if sex_level==`l' & mal_sbp_interact==0, lcols(sample) effect(OR) null(1) xlabel(-0.0001, -0.001, -0.01, 0.0) astext(50) dp(6) nograph
	
	local x = `x'+1
	local beta_s = (r(ES))
	local lci_s = (r(ci_low))
	local uci_s = (r(ci_upp))
	local p_val_s = r(p_z)	
	putexcel A`x'="SBP (SD)" B`x'="SBP" C`x'="`l' sex" D`x'="Difference" E`x'=`beta_s' F`x'=`lci_s' G`x'=`uci_s' H`x'=`p_val_s'
	
}
	
	
	*And now the interactive coefficients
	*This time we specify that we are interested in the interactions hence mal_sbp_interact==1
	metan log_or lci uci if  mal_sbp_interact==1 & sex_level==., nograph
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify
	
	*Specify the local for excel row number and to store result values
	local x = 3
	local beta_int = r(ES)
	local lci_int = r(ci_low)
	local uci_int = r(ci_upp)
	local p_value = r(p_z)

	*Store the results in the excel file
	putexcel A`x'="SBP combined P=5x10-8 (SD)" B`x'="SBP" C`x'="Interaction" D`x'="Additive" E`x'=`beta_int' F`x'=`lci_int' G`x'=`uci_int' I`x'=`p_value'
	
	*Then the interactive coefficient but sex stratified 
	levelsof sex_level, local(levels)
	local x = 7
	
	foreach l of local levels {
	
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify
	
	metan log_or lci uci if sex_level==`l' & mal_sbp_interact==1, lcols(sample) effect(OR) null(1) xlabel(-0.0001, -0.001, -0.01, 0.0) astext(50) dp(6) nograph
	
	local x = `x'+1
	local beta_s = (r(ES))
	local lci_s = (r(ci_low))
	local uci_s = (r(ci_upp))
	local p_val_s = r(p_z)	
	putexcel A`x'="SBP combined P=5x10-8 (SD)" B`x'="SBP" C`x'="`l' sex" D`x'="Additive" E`x'=`beta_s' F`x'=`lci_s' G`x'=`uci_s' I`x'=`p_val_s'
	
	
}
	*Now want to meta-analyse the maltreatment coeefficient too - only interaction results
	metan mal_beta mal_lci mal_uci if mal_sbp_interact==1 & sex_level==., nograph
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify

	*Specify the local for excel row number and to store result values
	local x = 4
	local mal_beta = r(ES)
	local mal_lci = r(ci_low)
	local mal_uci = r(ci_upp)
	local mal_p = r(p_z)
	
	*Store the results in the excel file
	putexcel A`x'="Maltreatment score" B`x'="SBP" C`x'="Interaction" D`x'="Additive" E`x'=`mal_beta' F`x'=`mal_lci' G`x'=`mal_uci' I`x'=`mal_p'

	*Then the maltreatment coefficient but sex stratified 
	levelsof sex_level, local(levels)
	local x = 9
	
	foreach l of local levels {
	
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify
	
	metan mal_beta mal_lci mal_uci if sex_level==`l', lcols(sample) effect(OR) null(1) xlabel(-0.0001, -0.001, -0.01, 0.0) astext(50) dp(6) nograph
	
	local x = `x'+1
	local beta_s = (r(ES))
	local lci_s = (r(ci_low))
	local uci_s = (r(ci_upp))
	local p_val_s = r(p_z)	
	putexcel A`x'="Maltreatment score" B`x'="SBP" C`x'="`l' sex" D`x'="Additive" E`x'=`beta_s' F`x'=`lci_s' G`x'=`uci_s' H`x'=`p_val_s'
	
	
}


	*And then the 3-way interaction 
	metan beta_3 lci_3 uci_3 if sex_int == 1, nograph 
	putexcel set filename, sheet(sd_grs_sbp_5e8) modify
	
	*Specify the local for excel row number and to store result values
	local x = 12
	local b_3 = r(ES)
	local l_3 = r(ci_low)
	local u_3 = r(ci_upp)
	local p_3 = r(p_z)
	
	*Store the results in the excel file
	putexcel  C`x'=" 3 way Interaction" D`x'="Additive" E`x'=`b_3' F`x'=`b_3' G`x'=`b_3' I`x'=`p_3'


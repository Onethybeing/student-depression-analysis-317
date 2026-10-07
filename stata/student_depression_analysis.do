/*==============================================================================
  Determinants of Depression among University Students
  Stata replication of notebooks/student_depression_analysis.ipynb

  Data   : data/student_sample_317.csv (317 students, ages 18-24,
           CGPA on a 10-point scale, all above 6.00)
  Outcome: depression (1 = Yes, 0 = No)
  Models : binary logit and probit with marginal effects, LPM and subsample
           robustness checks

  Run from the repository root:   do stata/student_depression_analysis.do
  Requires Stata 14 or newer (no user-written packages needed).
==============================================================================*/

clear all
set more off
version 14

capture mkdir "results"
capture mkdir "results/stata"
capture log close
log using "results/stata/student_depression_analysis.log", replace text

*------------------------------------------------------------------------------*
* 1. Load data and check quality
*------------------------------------------------------------------------------*
import delimited "data/student_sample_317.csv", varnames(1) clear

describe
misstable summarize
duplicates report gender age family_background relationship_status ///
    living_style current_semester cgpa friend_circle health_condition ///
    living_expense depression

*------------------------------------------------------------------------------*
* 2. Data transformation and variable coding
*   - Complicated relationship (n = 11, none depressed) merged with
*     "in a relationship": single = 1 vs not single = 0
*   - Living expense (Low < 10,000; Mid 10,000-14,999; High >= 15,000):
*     dummies for Mid and High, Low is the base group
*   - Health condition kept for descriptive analysis only; it separates the
*     outcome almost perfectly, so it enters only the robustness model (5)
*------------------------------------------------------------------------------*
rename depression depression_txt

gen byte depression    = depression_txt == "Yes"
gen byte male          = gender == "Male"
rename current_semester semester
gen byte fam_lowermid  = family_background == "Lower Middle class"
gen byte fam_rich      = family_background == "Rich class"
gen byte cost_mid      = living_expense == "Mid"
gen byte cost_high     = living_expense == "High"
gen byte single        = relationship_status == "Single"
gen byte live_family   = living_style == "Family"
gen byte has_friends   = friend_circle == "Yes"
gen byte health_poor   = inlist(health_condition, "Abnormal", "Critical")
gen byte health_level  = 1*(health_condition=="Good") + 2*(health_condition=="Normal") ///
                       + 3*(health_condition=="Abnormal") + 4*(health_condition=="Critical")

label variable depression   "Depression (1 = Yes)"
label variable male         "Male"
label variable age          "Age (years)"
label variable cgpa         "CGPA"
label variable semester     "Current semester"
label variable cost_mid     "Mid living expense"
label variable cost_high    "High living expense"
label variable fam_lowermid "Lower-middle class"
label variable fam_rich     "Rich class"
label variable single       "Single"
label variable live_family  "Lives with family"
label variable has_friends  "Has friend circle"
label variable health_poor  "Poor health"
label variable health_level "Health level (1-4)"

global X "i.male age cgpa semester i.cost_mid i.cost_high i.fam_lowermid i.fam_rich i.single i.live_family i.has_friends"
global XC "male age cgpa semester cost_mid cost_high fam_lowermid fam_rich single live_family has_friends"

*------------------------------------------------------------------------------*
* 3. Descriptive analysis
*------------------------------------------------------------------------------*
summarize depression $XC health_poor
tabulate depression
tabstat age cgpa semester, by(depression) statistics(mean sd n)

* Depression rate by group (row percentages) with chi-square and Cramer's V
foreach v in gender family_background relationship_status living_style ///
             living_expense friend_circle health_condition {
    display _newline as text "==== `v' ===="
    tabulate `v' depression_txt, row chi2 V
}

* Fisher exact test for the 2x2 tables
foreach v in gender living_style friend_circle {
    tabulate `v' depression_txt, exact
}

* Correlation matrix of coded variables
correlate depression $XC health_level

*------------------------------------------------------------------------------*
* 4. Difference in means: Welch t-test and Mann-Whitney (rank-sum) test
*------------------------------------------------------------------------------*
foreach v in cgpa age semester {
    ttest `v', by(depression) unequal
    ranksum `v', by(depression)
}

*------------------------------------------------------------------------------*
* 5. Multicollinearity check
*------------------------------------------------------------------------------*
quietly regress depression $XC
estat vif

*------------------------------------------------------------------------------*
* 6. Binary logit model
*------------------------------------------------------------------------------*
logit depression $X
estimates store m1_logit

* Odds ratios
logit depression $X, or

* Average marginal effects (discrete change for 0/1 variables)
estimates restore m1_logit
margins, dydx(*)
* Marginal effects at the means (equivalent of the older mfx command)
margins, dydx(*) atmeans

* Goodness of fit and diagnostics
estimates restore m1_logit
estat classification                      // cut-off 0.5
summarize depression, meanonly
estat classification, cutoff(`r(mean)')   // cut-off = sample share depressed
estat gof, group(10) table                // Hosmer-Lemeshow test
lroc, nograph                             // area under the ROC curve
estimates restore m1_logit
display "McFadden pseudo R2 = " e(r2_p) "   LR chi2(" e(df_m) ") = " e(chi2)
estat ic                                  // AIC and BIC

* Predicted probability by CGPA for single vs not single (others at means)
margins single, at(cgpa = (6.25(0.25)10)) atmeans
marginsplot, title("Predicted probability of depression") ytitle("Pr(depression)") name(profile, replace)
graph export "results/stata/predicted_probability_cgpa.png", replace

*------------------------------------------------------------------------------*
* 7. Binary probit model
*------------------------------------------------------------------------------*
probit depression $X
estimates store m3_probit
margins, dydx(*)
margins, dydx(*) atmeans
estat classification
estat gof, group(10)
lroc, nograph
estat ic

*------------------------------------------------------------------------------*
* 8. Robustness checks
*------------------------------------------------------------------------------*
* (2) Logit with robust standard errors
logit depression $X, vce(robust)
estimates store m2_logit_robust

* (4) Linear probability model with robust standard errors
regress depression $XC, vce(robust)
estimates store m4_lpm

* (5) Logit adding the poor-health indicator
logit depression $X i.health_poor
estimates store m5_logit_health
lrtest m1_logit m5_logit_health

* (6) and (7) Male and female students separately.
* A logit cannot be used: among male students none of the non-single students
* is depressed, so single predicts the outcome perfectly (Stata would drop it
* with "single != 0 predicts failure perfectly"). The LPM is used instead.
tabulate single depression if male == 1
regress depression age cgpa semester cost_mid cost_high fam_lowermid fam_rich ///
        single live_family has_friends if male == 1, vce(robust)
estimates store m6_male

regress depression age cgpa semester cost_mid cost_high fam_lowermid fam_rich ///
        single live_family has_friends if male == 0, vce(robust)
estimates store m7_female

* Side-by-side comparison of all models
estimates table m1_logit m2_logit_robust m3_probit m4_lpm m5_logit_health m6_male m7_female, ///
    b(%9.3f) se(%9.3f) stats(N r2_p r2 ll) varlabel

*------------------------------------------------------------------------------*
* 9. Save the coded data
*------------------------------------------------------------------------------*
save "results/stata/student_depression_coded.dta", replace
export excel using "results/stata/student_depression_coded.xlsx", firstrow(variables) replace

log close

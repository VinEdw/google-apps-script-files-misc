# google-apps-script-files-misc

This repository is a container for a miscellaneous assortment of single script projects done with *Google Apps Script*. In my mind, it did not make sense to create a repo for each one of them. Below is a short description for each one of them.

---

## chemical-equation-formatter

This script is bound to a *Google Docs*. When run, it takes the text highlighted by the user and turns any *chemical equation like* pattern and formats it accordingly. 

Examples: 

| Initial Text | Becomes... |
| --- | --- |
| BaCl2 (aq) + Na2SO4 (aq) → BaSO4 (s) + 2NaCl (aq) | BaCl<sub>2</sub> <sub>(aq)</sub> + Na<sub>2</sub>SO<sub>4</sub> <sub>(aq)</sub> → BaSO<sub>4</sub> <sub>(s)</sub> + 2NaCl <sub>(aq)</sub> |
| Ba 2+ (aq) + 2Cl - (aq) + 2 Na + (aq) + SO4 2- (aq) → BaSO4 (s) + 2Na + (aq) + 2Cl - (aq) | Ba <sup>2+</sup> <sub>(aq)</sub> + 2Cl <sup>-</sup> <sub>(aq)</sub> + 2 Na <sup>+</sup> <sub>(aq)</sub> + SO4 <sup>2-</sup> <sub>(aq)</sub> → BaSO<sub>4</sub> <sub>(s)</sub> + 2Na + <sub>(aq)</sub> + 2Cl <sup>-</sup> <sub>(aq)</sub> |
| Ba 2+ (aq) + SO4 2- (aq) → BaSO4 (s) | Ba <sup>2+</sup> <sub>(aq)</sub> + SO<sub>4</sub> <sup>2-</sup> <sub>(aq)</sub> → BaSO<sub>4</sub> <sub>(s)</sub> |

## make-folders-by-week

This script is unbound. When run, it creates numbered folders for each week from the start date to the end date. This was originally made to help a friend who works at an afterschool program prepare *Google Drive* folders for the year. 

## organize-availability-survey-results

This script is bound to a *Google Sheets*. When run, it takes the results for a survey and creates a sheet for each unique response for a given question. This sheet will have a `=QUERY()` formula added that will grab all the responses from the main response sheet (one linked to a *Google Form*) that have that specific answer choice. The `=QUERY()` will also sort the responses slightly. This was originally made to sort the student responses for a tutoring availability survey. The original add-on that was being used was breaking due to some answer choices having unescaped double quotes (" should go to "") and lengths exceeding the maximum allowed for sheet names (100 characters). 


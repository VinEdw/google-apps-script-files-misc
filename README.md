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

## hybrid-meeting-av-scheduling-tool

This script is bound to a *Google Sheets*. It has multiple functions that can be run to help with scheduling brothers who help out at the [meetings](https://www.jw.org/en/jehovahs-witnesses/meetings/video-kingdom-hall/) of Jehovah's Witnesses. Each person has the jobs they can do, as well as parts they have already been scheduled for. The goal is to schedule without any *conflicts*.

## make-folders-by-week

This script is unbound. When run, it creates numbered folders for each week from the start date to the end date. This was originally made to help a friend who works at an afterschool program prepare *Google Drive* folders for the year. 

## mastermind-game-pvp

This script is bound to a *Google Sheets*. It implements a multiplayer version of *Mastermind*, where up to 4 players can race to finish the same puzzle, or puzzles created for each other, first. The game is played within the [spreadsheet](https://docs.google.com/spreadsheets/d/1TbS8g0OFPeNBDlU8RlJDZjJ-mJ7ELOjlkFiQfrP_c7Y/edit?usp=sharing).

## mastermind-game-solver

This script is bound to a *Google Sheets*. It implements a solver for the game *Mastermind*. After inputting the guesses made so far and their result, the *best* next guess is proposed. The solver is run within the [spreadsheet](https://docs.google.com/spreadsheets/d/1ZpNXu9WKU0gVewmiPm0RZIYGGJAD7vWg9oVf_-eFhW4/edit?usp=sharing).

## meeting-part-colorizer

This script is bound to a *Google Docs*. When run, it looks at the whole document and turns and paragraph that starts with the pattern `{letter}:`, and makes the text have a specific color. As of now there is no driver function, so the function just uses `V` as the letter and <span style="color:cornflowerblue;">cornflowerblue</span> as the color. 

## organize-availability-survey-results

This script is bound to a *Google Sheets*. When run, it takes the results for a survey and creates a sheet for each unique response for a given question. This sheet will have a `=QUERY()` formula added that will grab all the responses from the main response sheet (one linked to a *Google Form*) that have that specific answer choice. The `=QUERY()` will also sort the responses slightly. This was originally made to sort the student responses for a tutoring availability survey. The original add-on that was being used was breaking due to some answer choices having unescaped double quotes (" should go to "") and lengths exceeding the maximum allowed for sheet names (100 characters). 


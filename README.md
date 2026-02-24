<a id="readme-top"></a>
<!-- PROJECT SHIELDS -->
<!--
*** I'm using markdown "reference style" links for readability.
*** Reference links are enclosed in brackets [ ] instead of parentheses ( ).
*** See the bottom of this document for the declaration of the reference variables
*** for contributors-url, forks-url, etc. This is an optional, concise syntax you may use.
*** https://www.markdownguide.org/basic-syntax/#reference-style-links
-->
[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![project_license][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]


<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/OlsonTyler0/Projection-Automation">
    <img src="images/logo.png" alt="Logo" width="360" height="160">
  </a>

<h3 align="center">Graduate Projection Automation</h3>

  <p align="center">
    Automates student course projections management for graduate programs using Excel and VBA.
    <br />
    <br />
    <a href="https://github.com/OlsonTyler0/Projection-Automation/issues/new?labels=bug">Report Bug</a>
    &middot;
    <a href="https://github.com/OlsonTyler0/Projection-Automation/issues/new?labels=enhancement">Request Feature</a>
  </p>
</div>


<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

This project was created to assist with the Missouri State University's Graduate Program Office with automating taking student projections and turning them into data regarding the coruses for each semester to help plan course capacity. 

This is a script intended to be used with excel sheets historically used by the graduate office but has been made pretty modular.

<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- USAGE EXAMPLES -->
## Usage

### Basic Workflow

1. **Prepare your data:**
   - Import your data into excel using `Data -> Get Data -> From File -> Excel -> Projections sheet` imported sheet must be titled `imported-data` otherwise it will error
   - Ensure columns are named: M#, Name, Fall 2026, Spring 2027, etc. _
   - Make a copy of the current excel sheet, always keep a backup! This will add data to your sheet!

2. **Import the VBA script into excel**
    1. Copy the raw data from [UpdateMustHave.vba](https://raw.githubusercontent.com/OlsonTyler0/Projection-Automation/refs/heads/main/UpdateMustHaveClasses.vba) to your clipboard with `CTRL + A` + `CTRL + C` 
    2. Open your Excel workbook in which you want to use this system
    3. Open the VBA Editor `Alt + F11`
    4. Click `Insert -> Module`


    <img width="295" height="181" alt="image" src="https://github.com/user-attachments/assets/6218b106-a2c9-4729-ae64-8f3246f8402a" />


    5. Paste the content wifrom `UpdateMustHaveClasses.vba` into a new module in your workbook with `CTRL + V`
    6. Save the content to the workbook `CTRL + S` _(you will see a warning about how it will save to the workbook, just hit the "save" button)_
    7. (Optional) Save the workbook as `.xlsm` (macro-enabled format)
      This step would only need to be done if the macro needs to STAY on the excel workbook.

2. **Run the import:**
    - Press `ALT + F8` to open the macro menu; select "SetupDashboard" OR "ImportProjections" to start the script and allow it to change data.
   - The script will:
     - Read all course sheets (ACC 711 - FA26, MGT 534 - SP26, etc.)
     - Match students from imported-data to course sheets
     - Update Must Have (Yes/No) based on graduation semester
     - Auto-populate "Graduating SEMESTER" notes
     - Flag students no longer projected

3. **Review results:**
   - Check the Dashboard for enrollment summaries
   - Review Notes column (D) for graduation dates

### Data Format Example

**imported-data sheet:**
```
M#          | Name        | Fall 2026    | Spring 2027 | Summer 2027
EX123456    | John Doe    | ACC 211      |             |
            |             | ITC 623      |             |
            |             | FIN 3432     | MGT 534     |
            |             |              |             |
            |             |              |             |
EX234567    | Jane Smith  | ACC 211      | MGT 534     |
```

**Course sheet (e.g., "ACC 711 - FA26"):**
```
M#          | Name        | Must Have (Yes/No) | Notes
EX123456    | John Doe    | No                 |
EX234567    | Jane Smith  | Yes                | Graduating FA26
```

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- Logic Breakdown -->
## Logic
This script looks at the following in order to function

Parses all workbook titles and precieves them as what the usual predicition sheet does: "Fall 26" -> Loops through all the sheets one by one parsing the class title from the sheet title and looking for that specified class in that specified semester -> Adds the student if a match is found -> Creates a dashboard


<!-- LICENSE -->
## License

Distributed under the MIT Lisence. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- CONTACT -->
## Contact

Tyler Olson - [@tyler-s-olson](https://linkedin.com/in/tyler-s-olson) - to329s@missouristate.edu

Project Link: [https://github.com/OlsonTyler0/Projection-Automation](https://github.com/OlsonTyler0/Projection-Automation)

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/OlsonTyler0/Projection-Automation.svg?style=for-the-badge
[contributors-url]: https://github.com/OlsonTyler0/Projection-Automation/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/OlsonTyler0/Projection-Automation.svg?style=for-the-badge
[forks-url]: https://github.com/OlsonTyler0/Projection-Automation/network/members
[stars-shield]: https://img.shields.io/github/stars/OlsonTyler0/Projection-Automation.svg?style=for-the-badge
[stars-url]: https://github.com/OlsonTyler0/Projection-Automation/stargazers
[issues-shield]: https://img.shields.io/github/issues/OlsonTyler0/Projection-Automation.svg?style=for-the-badge
[issues-url]: https://github.com/OlsonTyler0/Projection-Automation/issues
[license-shield]: https://img.shields.io/github/license/OlsonTyler0/Projection-Automation.svg?style=for-the-badge
[license-url]: https://github.com/OlsonTyler0/Projection-Automation/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/tyler-s-olson

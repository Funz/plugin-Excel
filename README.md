[![Build Status](https://travis-ci.org/Funz/plugin-Excel.png)](https://travis-ci.org/Funz/plugin-Excel)

# Funz plugin: Excel

This plugin is dedicated to launch Excel calculations from Funz.
It supports the following syntax and features:

  * Input
    * file type supported: '*.xlsx' or '*.xlsm'
    * parameter syntax: 
      * variable syntax: `$(...)`
      * formula syntax: `@{...}`
      * comment char: `#`
    * example input file: [sheet.xlsx](https://github.com/Funz/plugin-Excel/blob/master/src/main/samples/sheet.xlsx)
      * will identify input cell commented by `$VariableName` as variables
  * Output
    * file type supported: 'out.txt' (which is standard output stream)
    * read any commented cell `=ResultName`
    * example output file:
        ```
        Microsoft (R) Windows Script Host Version 5.8
        Copyright (C) Microsoft Corporation. All rights reserved.
        
        z=125
        ```
        * will return output:
          * z=125


![Analytics](https://ga-beacon.appspot.com/UA-109580-20/plugin-Excel)

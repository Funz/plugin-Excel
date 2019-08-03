[![Build Status](https://travis-ci.org/Funz/plugin-MyPlugin.png)](https://travis-ci.org/Funz/plugin-MyPlugin)

___This repositry is intended to be forked as a basis for an new plugin implementation.___
___You can search for "MyPlugin" as a key to replace everywhere in this directory.___

# Funz plugin: MyPlugin

This plugin is dedicated to launch MyPlugin calculations from Funz.
It supports the following syntax and features:

  * Input
    * file type supported: '*.MyPlugin', any other format for resources
    * parameter syntax: 
      * variable syntax: `$(...)`
      * formula syntax: `@{...}`
      * comment char: `#`
    * example input file:
        ```
        ...
        ... $(x1~[1,2]) ...
        ... $x1 ...
        ...
        ```
      * will identify input:
        * x1, expected to vary inside [1,2]
        * x2, expected to vary inside [0,1] (by default)
      * replace `!{?x2 + 1.23 | #.###}` expression by its evaluation

  * Output
    * file type supported: '*.MyPlugin.out'
    * read any named value printed with `=`, like `...`
    * example output file:
        ```
        ...
        z= ...
        ...
        ```
        * will return output:
          * z=... 


![Analytics](https://ga-beacon.appspot.com/UA-109580-20/plugin-MyPlugin)

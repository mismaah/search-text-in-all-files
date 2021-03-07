### Install Dependencies

Before running script, install doc2txt and pptx

`pip install doc2txt python-pptx`

### Usage

Copy script.py inside the folder where you want to search text and run it.

Currently supports docx, txt and pptx file search.

The script also searches within all child folders recursively.

### Config

The CONTEXT_RANGE constant can be changed to specify how many words ahead or
beyond the matched word is shown in contexts.

The MATCH_OVER_SIMILARITY determines the lower limit of similarities of the 
matches results that will be shown.

### Examples

![Example 1](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex1.PNG?raw=true)
![Example 2](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex2.PNG?raw=true)
![Example 3](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex3.PNG?raw=true)

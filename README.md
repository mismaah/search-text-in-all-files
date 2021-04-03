### Install Dependencies

Before running scripts, install doc2txt, pptx and pdfminer.

`pip install doc2txt python-pptx pdfminer`

### Usage

Copy either script inside the folder where you want to search text and run it.

Currently supports docx, txt and pptx file search.

The scripts also searches within all child folders recursively.

trigram_token_match.py does a trigram similarity search where similar matches
are also found in addition to exact matches. Only works for single words. 

regex_match.py does a simple regular expression search. Can match any string and
is not limited to single words.

### Config

The CONTEXT_RANGE constant can be changed to specify how many words ahead or
beyond the matched word is shown in contexts.

The MATCH_OVER_SIMILARITY in trigram_token_match.py determines the lower limit
of similarities of the matches results that will be shown.

### Examples

#### trigram_token_match.py
![Example 1](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex1.PNG?raw=true)
![Example 2](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex2.PNG?raw=true)
![Example 3](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex3.PNG?raw=true)

#### regex_match.py
![Example 1](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex4.PNG?raw=true)
![Example 2](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex5.PNG?raw=true)
![Example 3](https://github.com/mismaah/search-text-in-all-files/blob/main/examples/ex6.PNG?raw=true)

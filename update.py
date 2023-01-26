pending = "We have not had a chance to write a definition yet, sorry for the inconvenience."

# https://www.mathjax.org/#gettingstarted
latex = '''<script src="https://polyfill.io/v3/polyfill.min.js?features=es6"></script>
<script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>'''

spacer = '\n<hr noshade>\n'
punct = ' .,:;-()s' # plurals are okay to link

def anchor(text, keyword, link):
    start = text.index(keyword)
    if text[start - 1] == '>':
        return text # already linked
    if text[start - 1] not in punct: 
        return text # do not link suffixes   
    end = text.index(keyword) + len(keyword)
    if end < len(text) and text[end] not in punct: # not a whole word
        return text    
    return text.replace(keyword, link)

from re import sub

def linkify(text):
  camel = sub(r'(_|-)+', ' ', text).title().replace(' ', '')
  return ''.join([ camel[0].lower(), text[1:] ])

def include(term, name, definition, references, appearances):
    output = f'<h2><a name="{name}">{term}</a></h2>\n'
    definition = definition.replace('\n\n', '</p><p>')# paragraph change
    output += f'<p>{definition}</p>\n\n'
    if len(references) > 0:
        output += f'<h3>References</h3>\n<ol>'
        for ref in references.split():
            url = ref.strip().lstrip()
            if 'http' in url: 
                output += f'<li><a href="{url}">{url}</a></li>'
            else: # not a proper URL, cannot hyperlink it
                output += f'<li>{url}</li>'
        output += '</ol>\n\n'
    (introduced, extended, used) = appearances
    li = len(introduced)
    le = len(extended)
    lu = len(used)
    if li > 0 or le > 0 or lu > 0:
        output += '<h3>Related courses</h3>\n'
        if li > 0:
            intro = sorted(list(introduced))
            output += 'This concept is introduced in ' + ' '.join(intro) + '.'
        if le > 0:
            ext = sorted(list(extended))
            output += 'This concept is revisited in more depth in ' + ' '.join(ext) + '.'
        if lu > 0:
            uses = sorted(list(used))
            output += 'This concept is employed in ' + ' '.join(uses) + '.'
    return output

import pandas as pd

terms = {}
references = {}
usage = {}
link = {}
name = {}
target = {}

glossary = pd.ExcelFile('Glossary.xlsx')
for sheet in glossary.sheet_names:
    data = glossary.parse(sheet)
    link[sheet] = dict()
    header = data.columns.values.tolist()
    definition = header.index('Definition')
    courses = header[1 : definition]
    for index, values in data.iterrows():
        term = str(values['Concept'])
        if term == 'nan':
            continue # thank you, excel
        intro = set()
        use = set()
        extend = set()
        for course in courses:
            if values[course] == 'intro': # introduced in this course
                intro.add(course)
            elif values[course] == 'extend': # more details provided in this course
                extend.add(course)
            elif values[course] == 'use': # used in this course, no further detail given
                use.add(course)
        concept = str(values['Concept'])
        name[concept] = linkify(concept)
        link[sheet][concept] = f'<a href="{sheet}.html#{name[concept]}">{concept}</a>'
        written = str(values['Definition'])
        refs = str(values['References'])
        refs = '' if refs == 'nan' else refs
        if written == 'nan':
            written = pending
        if concept not in terms: # new concept
            terms[concept] = written
            references[concept] = refs 
            usage[concept] = (intro, extend, use)            
        else: # not the first occurence of this term
            if terms[concept] != pending: # no need to repeat the apology
                terms[concept] += '\n\n' + written # paragraph change
            if len(refs) > 0:
                references[concept] += '\n' + refs
            (i, e, u) = usage[concept]
            usage[concept] = ( intro | i, extend | e, use | u)
    for term in terms:
        definition = terms[term]
        for other in terms:
            if other == term:
                continue # no self-referencing
            if other in definition:            
                if other in link[sheet]: # it is on this sheet
                    definition = anchor(definition, other, link[sheet][other])
                else: # check if it is on another sheet
                    for alt in link:
                        if other in link[alt]:
                            definition = anchor(definition, other, link[alt][other])
                            break # use the first match
        terms[term] = definition

for theme in link:
    with open(f'{theme}.html', 'w') as target:
        start = f'<!DOCTYPE html><html><head><title>{theme}</title>{latex}</head>'
        start += f'<body><h1>Glossary for {theme}</h1>\n'
        print(start, file = target)
        listing = sorted(list(link[theme].keys()))
        for term in listing:
            print(spacer, file = target)
            print(include(term, name[term], terms[term], references[term], usage[term]), file = target)
        print(spacer, file = target)
        end = '</body></html>'
        print(end, file = target)        

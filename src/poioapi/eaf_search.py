# -*- coding: utf-8 -*-
import poioapi.annotationgraph

'''
EAF_FILEPATH = "Luka.eaf"
EAF1_FILEPATH = "Luka1.eaf"
OUTPUT_XLS_PATH = "out.xls"
'''

# returns True if something was found
def eaf_search(input_path, sheet, search_query):
    annotation_graph = poioapi.annotationgraph.AnnotationGraph.from_elan(input_path)
    tree = annotation_graph.as_tree()
    tree.filter(search_query)
    wrote_something = tree.print_to_xls_sheet(sheet, filtered=True)
    return wrote_something

'''

query = dict()
query['type'] = 'and'
query['value'] = list()
query['tiers'] = ['word']

first_and = dict()
first_and['type'] = 'substr'
first_and['value'] = 'Ангелъ'
first_and['tiers'] = ['transcription']

second_and = dict()
second_and['type'] = 'substr'
second_and['value'] = 'Ангел'

query['value'].append(first_and)
query['value'].append(second_and)

print(eaf_search(EAF_FILEPATH, OUTPUT_XLS_PATH, query, "new_sheet"))
print(eaf_search(EAF1_FILEPATH, OUTPUT_XLS_PATH, query, "new_sheet1"))
'''
# -*- coding: utf-8 -*-

import ast

import os
import os.path

import pprint
import sys


import poioapi.annotationgraph


'''
EAF_FILEPATH = "Luka.eaf"
EAF1_FILEPATH = "Luka1.eaf"
OUTPUT_XLS_PATH = "out.xls"
'''

def eaf_search(
    input_path,
    output_path,
    search_query,
    sheet_name,
    result_list = None):
    """
    input_path: input stream or path to the input file.
    outpath_path: path to the output file, possibly non-existing.

    Returns True if found something.
    """

    try:
        annotation_graph = poioapi.annotationgraph.AnnotationGraph.from_elan(input_path)

    except:
        return False

    tree = annotation_graph.as_tree()
    tree.filter(search_query)

    wrote_something = (

        tree.print_to_xls(
            output_path,
            filtered = True,
            name = sheet_name,
            result_list = result_list))

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


# If we are launched as a script, for testing.

if __name__ == '__main__':

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    search_query = ast.literal_eval(sys.argv[3])
    sheet_name = sys.argv[4]

    if os.path.exists(output_path):
        os.remove(output_path)

    result_list = []

    result = (
        eaf_search(
            input_path,
            output_path,
            search_query,
            sheet_name,
            result_list))

    print(
        'eaf_search(\n  {0},\n  {1},\n  {2},\n  {3},\n  result_list)'.format(
            repr(input_path),
            repr(output_path),
            search_query,
            repr(sheet_name)))

    print(
        'result: {0}\nresult_list:\n{1}'.format(
            result,
            pprint.pformat(result_list, width = 144)))


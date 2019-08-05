# -*- coding: utf-8 -*-
import poioapi.annotationgraph

EAF_FILEPATH = "Luka.eaf"
OUTPUT_XLS_PATH = "out.xls"


def eaf_search(input_path, output_path, search_query):
    annotation_graph = poioapi.annotationgraph.AnnotationGraph.from_elan(input_path)
    tree = annotation_graph.as_tree()
    tree.filter(search_query)
    tree.print_to_xls(output_path, filtered=True)


query = list()
first_or = list()
second_or = list()
first_or.append("Ангелъ")
second_or.append("возвратиться")
second_or.append("мурдасть")
query.append(first_or)
query.append(second_or)

eaf_search(EAF_FILEPATH, OUTPUT_XLS_PATH, query)

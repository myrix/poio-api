# -*- coding: utf-8 -*-
#
# Poio Tools for Linguists
#
# Copyright (C) 2009-2013 Poio Project
# Author: Peter Bouda <pbouda@cidles.eu>
# URL: <http://media.cidles.eu/poio/>
# For license information, see LICENSE.TXT

import sys
import os
import codecs
import re
import glob
import subprocess
import itertools
import io

import poioapi.annotationgraph
import poioapi.data
import poioapi.io.graf

import graf


obt_tagger_command = \
    "/Users/pbouda/Projects/git-github/The-Oslo-Bergen-Tagger/tag-nostat-bm.sh"

ner_tag_for_type = {
    "by_navn" : "city",
    "etternavn" : "family_name",
    "fylkesnavn" : "county",
    "guttenavn" : "boy_name",
    "jentenavn" : "girl_name",
    "kommune_navn" : "district",
    "land_navn" : "country"
}

re_ner_type = re.compile("([^.]*)\.txt$")

def parse_ner_list(filename):
    mwes = set()
    with codecs.open(filename, "r", "utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("[ STEM <"):
                words = line[9:-3]
                if words.startswith('"'):
                    words = words[1:-1]
                words = re.split('", "', words)
                words = [w.lower() for w in words if w.lower() not in \
                    ['de', 'den', 'det']]
                mwes.add(tuple(words))
    return mwes

def obt_tagger(inputfile):
    proc = subprocess.Popen([obt_tagger_command, inputfile],
        cwd=os.path.dirname(obt_tagger_command), stdout=subprocess.PIPE)
    (out, err) = proc.communicate()
    return out.decode("utf-8")

def last_used_id_in_graf(graf):
    last_used_id = 0
    for n in graf.nodes:
        n_id = int(re.search("\d+$",n.id).group(0))
        if n_id > last_used_id:
            last_used_id = n_id
    return last_used_id

def add_ner_node(graf_graph, parent_nodes, ner_tag, last_used_id):
    node_id = "named_entity..na{0}".format(last_used_id)
    ner_node = graf.Node(node_id)
    ann = graf.Annotation("named_entity",
        {"annotation_value": ner_tag}, "a{0}".format(last_used_id))
    ner_node.annotations.add(ann)
    graf_graph.nodes.add(ner_node)
    for n in parent_nodes:
        graf_graph.create_edge(n, ner_node)
    return last_used_id + 1

def tags_for_tuple(ner_dict, variant_tuple):
    tags = set()
    for ner_tag in ner_dict:
        if variant_tuple in ner_dict[ner_tag]:
            tags.add(ner_tag)
    return tags


def main(argv):
    if len(argv) != 3:
        print("Usage: obt_ner_tagger.py inputfile outputfile")
        sys.exit(1)

    inputfile = os.path.realpath(argv[1])
    outputfile = os.path.realpath(argv[2])
    scriptpath = os.path.dirname(os.path.realpath(__file__))
    
    ner_dict = dict()

    for ner_file in glob.glob(os.path.join(scriptpath, "ner_lists", "*.txt")):
        ner_type = re_ner_type.search(ner_file).group(1)
        ner_set = parse_ner_list(ner_file)
        ner_dict[ner_tag_for_type[ner_type]] = ner_set

    # find out what the longest MWE is
    max_ngram = 0
    for ner_tag in ner_dict:
        for mwe in ner_dict[ner_tag]:
            if len(mwe) > max_ngram:
                max_ngram = len(mwe)

    # get the OBT tagger output
    obt_out = obt_tagger(inputfile)
    # Create an empty annotation graph
    ag = poioapi.annotationgraph.AnnotationGraph(None)
    ag.from_obt(io.StringIO(obt_out))
    hierarchy = ag.tier_hierarchies[0]
    hierarchy[1].append('named_entity') 
    ag.structure_type_handler = poioapi.data.DataStructureType(hierarchy)
    last_used_id = last_used_id_in_graf(ag.graf)

    for phrase in ag.root_nodes():
        mwe_nodes = list()
        mwe_words = list()
        for word in ag.nodes_for_tier("word", phrase):
            variants = list()
            is_prop = False
            for variant in ag.nodes_for_tier("variant", word):
                for tag in ag.nodes_for_tier("tag", variant):
                    # are we in a name?
                    if ag.annotation_value_for_node(tag) == "prop":
                        is_prop = True
                        variants.append(
                            ag.annotation_value_for_node(variant).lower())

            if is_prop:
                mwe_words.append(variants)
                mwe_nodes.append(word)

            # did we leave the prop/MWE?
            elif len(mwe_words) > 0:
                found_mwe = False
                tags = set()
                # check if this is a MWE
                for variant in itertools.product(*mwe_words):
                    tags = tags.union(tags_for_tuple(ner_dict, tuple(variant)))

                for t in tags:
                    last_used_id = add_ner_node(ag.graf, mwe_nodes,
                        ner_tag, last_used_id)
                    found_mwe = True

                if not found_mwe:
                    for i, word in enumerate(mwe_words):
                        tags = set()
                        found_ner = False
                        for variant in word:
                            tags = tags.union(
                                tags_for_tuple(ner_dict, tuple([variant])))
    
                        for t in tags:
                            last_used_id = add_ner_node(ag.graf, [mwe_nodes[i]],
                                ner_tag, last_used_id)
                            found_ner = True

                        if not found_ner:
                            # set default tag "proper_name"
                            last_used_id = add_ner_node(ag.graf,
                                [mwe_nodes[i]], "proper", last_used_id)

                mwe_words = list()
                mwe_nodes = list()


    writer = poioapi.io.graf.Writer()
    writer.write(outputfile, ag)


if __name__ == "__main__":
    main(sys.argv)
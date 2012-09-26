# -*- coding: utf-8 -*-
#
# Poio Tools for Linguists
#
# Copyright (C) 2009-2012 Poio Project
# Author: António Lopes <alopes@cidles.eu>
# URL: <http://www.cidles.eu/ltll/poio>
# For license information, see LICENSE.TXT
""" This module is to create the regions
of the WFW one by one
"""
from xml.dom.minidom import Document
from poioapi import annotationtree
from poioapi import data

import os
import pickle
import codecs

class CreateWfwFile:

    def __init__(self, filepath):
        self.filepath = filepath

    def create_wfw_xml(self):

        # Initialize the variable
        annotation_tree = annotationtree.AnnotationTree(data.GLOSS)

        # Open the file
        file = open(self.filepath, "rb")
        annotation_tree.tree = pickle.load(file)

        doc = Document()
        graph = doc.createElement("graph")
        graph.setAttribute("xmlns:graf", "http://www.xces.org/ns/GrAF/1.0/")
        doc.appendChild(graph)

        # Header
        graphheader = doc.createElement("graphHeader")
        labelsdecl = doc.createElement("labelsDecl")
        graphheader.appendChild(labelsdecl)
        graph.appendChild(graphheader)

        dependencies = doc.createElement('dependencies')
        dependson = doc.createElement('dependsOn')
        dependson.setAttribute('f.id','word')
        dependencies.appendChild(dependson)
        graphheader.appendChild(dependencies)

        ann_spaces = doc.createElement('annotationSpaces')
        ann_space = doc.createElement('annotationSpace')
        ann_space.setAttribute('as.id','wfw')
        ann_spaces.appendChild(ann_space)
        graphheader.appendChild(ann_spaces)

        # Start XML file
        basename = self.filepath.split('.pickle')
        file = os.path.abspath(basename[0] + '-wfw.xml')
        f = codecs.open(file,'w','utf-8')

        id_counter = 0

        # Verify the elements
        for element in annotation_tree.elements():

            # Get the utterance
            utterance = element[1]

            for wfw_el in utterance:
                for el in wfw_el[1]:
                    st = el[1].get('annotation')

                    if (st == ''):
                        id_counter+=1
                        continue

                    # Creating the node with link
                    node = doc.createElement("node")
                    node.setAttribute("xml:id", "wfw-n"
                    + str(id_counter)) # Node number

                    # Creating the node
                    link = doc.createElement("link")
                    link.setAttribute("targets", "word-r"
                    + str(id_counter)) # ref
                    node.appendChild(link)

                    graph.appendChild(node)

                    # Creating the features and the linkage
                    a = doc.createElement("a")
                    a.setAttribute("xml:id", "wfw-"
                    + str(id_counter)) # id
                    a.setAttribute("label", "wfw") # label
                    a.setAttribute("ref", "wfw-n"
                    + str(id_counter)) # ref
                    a.setAttribute("as", "wfw") # as

                    # Feature structure
                    feature_st = doc.createElement("fs")
                    feature = doc.createElement("f")
                    feature.setAttribute("name","wfw")
                    feature.setAttribute("value",st)
                    value = doc.createTextNode(st) # Value
                    feature.appendChild(value)
                    feature_st.appendChild(feature)

                    a.appendChild(feature_st)
                    graph.appendChild(a)

                    id_counter+=1

        # Write the content in XML file
        f.write(doc.toprettyxml(indent="  "))

        #Close XML file
        f.close()
# -*- coding: utf-8 -*-
#
# Poio Tools for Linguists
#
# Copyright (C) 2009-2013 Poio Project
# Author: António Lopes <alopes@cidles.eu>
# URL: <http://media.cidles.eu/poio/>
# For license information, see LICENSE.TXT
"""This module contains classes to access to
Typecraf files.
"""

import os
import re

from xml.dom import minidom
from xml.etree.ElementTree import Element, SubElement, tostring
import xml.etree.ElementTree as ET

from poioapi.io import header
from poioapi.io.graf import GrAFWriter

from graf.graphs import Graph
from graf import Node, Edge
from graf import Annotation, AnnotationSpace
from graf import Region

class Typecraft:
    """
    Class that will handle the parse of
    Typecraft files.

    """

    def __init__(self, filepath):
        """Class's constructor.

        Parameters
        ----------
        filepath : str
            Path of the elan file.

        """

        self.filename = os.path.basename(filepath)
        self.filepath = filepath
        (self.basedirname, _) = os.path.splitext(os.path.abspath(self.filepath))

        self.xml_files_map = {}

        self.graf = GrAFWriter()

    def typecraft_to_elan(self):
        tree = ET.parse(self.filepath)
        root = tree.getroot()

        xml_namespace = re.search('\{(.*)\}', root.tag).group()

        phrases = root.findall(xml_namespace+'phrase')

        graph = Graph()

        globaltags = xml_namespace+"globaltags"
        word = xml_namespace+"word"
        words_number = 1
        morpheme_number = 1
        gloss_number = 1

        element_tree = self.graf.create_xml_graph_header('phrase', None)
        word_element_tree = self.graf.create_xml_graph_header('word', 'phrase')
        morph_element_tree = self.graf.create_xml_graph_header('morpheme','word')
        gloss_element_tree = self.graf.create_xml_graph_header('gloss','morpheme')

        for phrase in phrases:
            index = phrase.attrib['id']

            node_id = "phrase"+"/n"+index
            node = Node(node_id)

            anchors = ['10','100']

            region_id = "phrase"+"/r"+index
            region = Region(region_id, *anchors)

            node.add_region(region)

            # Annotation
            annotation = Annotation("phrase", None,
                "phrase/a"+index)

            for attribute in phrase.attrib:
                annotation.features[attribute] = phrase.attrib[attribute]

            from_node = node

            for elements in phrase:
                if elements.tag == word:
                    word_node_id = "word"+"/n"+str(words_number)
                    word_node = Node(word_node_id)

                    word_anchors = ['1','9']

                    word_region_id = "word"+"/r"+str(words_number)
                    word_region = Region(word_region_id, *word_anchors)

                    word_node.add_region(word_region)

                    word_edge_id = "word/e"+str(words_number)
                    word_edge = Edge(word_edge_id, from_node, word_node)

                    graph.edges.add(word_edge)

                    # Annotation
                    word_ann = Annotation("word",
                        None,"word/a"+str(words_number))

                    for word_key in elements.attrib:
                        word_ann.features[word_key] = elements.attrib[word_key]

                    # Look for the morpheme
                    for word_elements in elements:
                        key = str(word_elements.tag).split(xml_namespace)
                        if key[1] == "pos":
                            word_ann.features[key[1]] = word_elements.text
                        elif key[1] == "morpheme":
                            morph_node_id = "morpheme"+"/n"+str(morpheme_number)
                            morph_node = Node(morph_node_id)

                            morph_edge_id = "morpheme/e"+str(morpheme_number)
                            morph_edge = Edge(morph_edge_id, word_node, word_node)

                            graph.edges.add(morph_edge)

                            # Annotation
                            morph_ann = Annotation("morpheme",
                                None,"morpheme/a"+str(morpheme_number))

                            for morph_key in word_elements.attrib:
                                morph_ann.features[morph_key] = word_elements.attrib[morph_key]

                            morph_node.annotations.add(morph_ann)
                            graph.nodes.add(morph_node)

                            for gloss in word_elements:
                                gloss_ann = Annotation("gloss", None,
                                    "gloss/a"+str(gloss_number))
                                gloss_ann.features['annotation_value'] = gloss.text

                                graph.nodes[morph_node_id].annotations.add(gloss_ann)

                                annotation_space = AnnotationSpace('gloss')
                                annotation_space.add(gloss_ann)
                                graph.annotation_spaces.add(annotation_space)

                                gloss_element_tree = self.graf.create_node_annotation(gloss_element_tree,
                                    gloss_ann, morph_node_id)

                                self.xml_files_map['gloss'] = gloss_element_tree

                                gloss_number+=1

                            annotation_space = AnnotationSpace('morpheme')
                            annotation_space.add(morph_ann)
                            graph.annotation_spaces.add(annotation_space)

                            morph_element_tree = self.graf.create_node_with_region(morph_element_tree,
                                morph_ann, morph_node.id, morph_node, None, None,
                                word_node, morph_edge)

                            self.xml_files_map['morpheme'] = morph_element_tree

                            morpheme_number+=1

                    word_node.annotations.add(word_ann)

                    annotation_space = AnnotationSpace('word')
                    annotation_space.add(word_ann)

                    graph.nodes.add(word_node)
                    graph.regions.add(word_region)
                    graph.annotation_spaces.add(annotation_space)

                    word_element_tree = self.graf.create_node_with_region(word_element_tree,
                        word_ann, word_node.id, word_node, word_region, word_anchors,
                        from_node, word_edge)

                    self.xml_files_map['word'] = word_element_tree

                    words_number+=1
                elif elements.tag == globaltags:
                    print(globaltags)
                else:
                    key = str(elements.tag).split(xml_namespace)
                    annotation.features[key[1]] = elements.text

            node.annotations.add(annotation)

            annotation_space = AnnotationSpace('phrase')
            annotation_space.add(annotation)

            graph.nodes.add(node)
            graph.regions.add(region)
            graph.annotation_spaces.add(annotation_space)

            element_tree = self.graf.create_node_with_region(element_tree,
                annotation, node_id, node, region, anchors,
                None, None)

            self.xml_files_map['phrase'] = element_tree

        self.generate_graf_files()

    def generate_graf_files(self):
        """This method will create the GrAF Xml files.
        But first is need to create the GrAF object in
        order to get the values.
        This method will also create the header file.

        """

        self.header = header.HeaderFile(self.basedirname)
        self.header.filename = 'oo'
        self.header.primaryfile = 'oo'
        self.header.dataType = 'text' # Type of the origin data file

        for elements in self.xml_files_map.items():
            file_name = elements[0]
            extension = file_name+".xml"
            filepath = self.basedirname+"-"+extension
            loc = os.path.basename(filepath)
            self.header.add_annotation(loc, file_name)
            file = open(filepath,'wb')
            element_tree = elements[1]
            doc = minidom.parseString(tostring(element_tree))
            file.write(doc.toprettyxml(indent='  ', encoding='utf-8'))
            file.close()

        self.header.create_header()
#!/usr/bin/env python

"""
This file will take arguments which will fill in a programme for working. 
"""
import argparse
from docx import *
import sys



if __name__ == '__main__':
   
   parser = argparse.ArgumentParser(description = 'Create a word document for a school booking.', epilog="(c)2014 A Cripps")
   parser.add_argument('--school','-s', dest='school', metavar='S', help="A string of the name of the School.", required=True)
   parser.add_argument('--file', '-f', metavar='F', help="Filename to be output.")


   args = parser.parse_args()
   

   print args.school
   # Default set of relationships - the minimum components of a document
   relationships = relationshiplist()

   # Make a new document tree - this is the main part of a word document 
   document = newdocument()

   # This xpath location is where most interesting content lives
   body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

   # Append two headings in to a paragraph
   
   body.append(heading(args.school, 1))




   # Create our properties, contenttypes, and other support files
   title    = 'Python docx demo'
   subject  = 'A practical example of making docx from Python'
   creator  = 'Mike MacCana'
   keywords = ['python', 'Office Open XML', 'Word']



   coreprops = coreproperties(title=title, subject=subject, creator=creator, keywords=keywords)
   appprops = appproperties()
   contenttypes = contenttypes()
   websettings = websettings()
   wordrelationships = wordrelationships(relationships)
   if (args.file):
      print "file is:" + args.file
      filename = args.file+".docx"
   else:
      filename = args.school+".docx"
   savedocx(document, coreprops, appprops, contenttypes, websettings, wordrelationships,filename)

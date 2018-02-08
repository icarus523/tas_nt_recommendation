#!/usr/bin/env python

import sys
import json
import os
from pprint import pprint
from PyRTF import *

def importdata(filename):
        with open(filename) as data_file:
                data = json.load(data_file)

        #	pprint(data)
        return data

def tas_nt(filename) :
        # filename = "tas_nt_recommendation_66661.json"
        doc     = Document()
        ss      = doc.StyleSheet
        section = Section()
        doc.Sections.append( section )

        thin_edge  = BorderPS( width=20, style=BorderPS.SINGLE )
        thin_frame  = FramePS( thin_edge,  thin_edge,  thin_edge,  thin_edge )

        p = Paragraph( ss.ParagraphStyles.Heading1 )
        p.append('Attachment 1')
        section.append( p )

        gamesoftware_table = Table(2000, 2000, 2000, 3500)
        c1 = Cell( Paragraph( ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'Game Software'), thin_frame, span=4) 
        gamesoftware_table.AddRow( c1)

        c2 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'Game Name'), thin_frame)
        c3 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'Game Label'), thin_frame)
        c4 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'Base Version'), thin_frame)
        c5 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'HMAC SHA1 Hash (Seed 0)'), thin_frame)
        gamesoftware_table.AddRow(c2, c3 ,c4, c5)

        jsondata = importdata(filename) # read JSON file
        gamename = str(jsondata["Game_Name"])
        gamelabel = str(jsondata["Game_Label"])
        baseversion = str(jsondata["Base_Version"])
        hmacsha1 = str(jsondata["HMACSHA1"])

        c1 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, gamename), thin_frame)
        c2 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, gamelabel), thin_frame)
        c3 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, baseversion), thin_frame)
        c4 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, TEXT(hmacsha1, font=ss.Fonts.CourierNew)), thin_frame)
        gamesoftware_table.AddRow(c1, c2, c3, c4)

        section.append( gamesoftware_table )

        p = Paragraph( ss.ParagraphStyles.Normal )
        p.append(' ')
        section.append( p )

        psdinformation_table = Table(9500/3, 9500/3, 9500/3)
        c1 = Cell( Paragraph( ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'PSD Information'), thin_frame, span=3) 

        psdinformation_table.AddRow(c1)

        c2 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'PSD Name / Label'), thin_frame )
        c3 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'Position'), thin_frame)
        c4 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Heading2, 'HMAC SHA1 Hash (Seed 0)'), thin_frame )
        psdinformation_table.AddRow(c2, c3, c4)

        #print jsondata["PSD_Info"]
        # json_object[0].items()
        # json_object[0]["title"]

        psdname = ''
        psd_position = ''
        psd_SHA1 = ''
        
        for item in jsondata["PSD_Info"]:
                for attribute, value in item.iteritems():
                        if attribute == 'PSD_Label':
                                psdname = str(value)
                        elif attribute == 'PSD_Position':
                                psd_position = str(value)
                        elif attribute == 'PSD_Hash':
                                psd_SHA1 = str(value)

                #psdname = str(jsondata["PSD_Info"]["PSD_Label"])
                #psd_position = str(jsondata["PSD_Info"]["PSD_Position"])
                #psd_SHA1 = str(jsondata["PSD_Info"]["PSD_Hash"])

                c1 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, psdname), thin_frame)
                c2 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, psd_position), thin_frame)
                c3 = Cell( Paragraph(ParagraphPS(alignment = ParagraphPS.CENTER), ss.ParagraphStyles.Normal, TEXT(psd_SHA1, font=ss.Fonts.CourierNew)), thin_frame)
                psdinformation_table.AddRow(c1, c2, c3)

        section.append( psdinformation_table )

        return doc

def OpenFile( name ) :
	return file( '%s.rtf' % name, 'w' )

def main(): 
        DR = Renderer()

        if len(sys.argv) < 2:
                print "Usage: python tas_nt_recommendation_letter.py <args1> (json input file)"
                sys.exit(2)

        filename = sys.argv[1]
        if os.path.isfile(filename) :      
                doc1 = tas_nt(filename)
                DR.Write(doc1, OpenFile(filename.lstrip("tas_nt_recommendation").rstrip(".json")))
                print "Finished"
        else:
                print filename + ": file can't be found, check filename input."

if __name__ == '__main__' : main()

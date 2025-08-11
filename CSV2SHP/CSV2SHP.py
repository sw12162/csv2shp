#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on 25/06/2025 08:12

@author: sw12162
"""

# Splash screen import
try:
    import pyi_splash
except:
    pass

# Externally installed packages
from PyQt6.QtWidgets import QWidget, QPushButton, QLabel, QLineEdit, QComboBox, QFileDialog, QGridLayout,  QApplication, QMessageBox
from PyQt6.QtGui import QIcon
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point, LineString
from sklearn.neighbors import NearestNeighbors
import networkx as nx
import numpy as np
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.figure import Figure
import contextily as ctx
from sklearn.cluster import DBSCAN
from networkx.utils import UnionFind



# Base python packages
from collections import Counter, defaultdict
from traceback import format_exc
from pathlib import Path
import os
import re
import csv
import ctypes


#### START ####
myappid = 'BKK.CSV2SHP.v1.6'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

# Close splash on startup
try:
    pyi_splash.close()
except: 
    pass

# Create a window
class InputDialog(QWidget):
    def __init__(self):       
        super(InputDialog,self).__init__()
        
        # GUI labels
        in_name_lab = QLabel("Input file:")
        out_name_lab = QLabel("Output folder:")
        num_lab = QLabel("Måler og prosjekt nummer")
        ptema_lab = QLabel("PTema:")
        meth_lab = QLabel("Målemetode:")
        noyakt_lab = QLabel("Noyaktighet:")
        syn_lab = QLabel("Synbarhet:")
        trase_lab = QLabel("Auto-trase:")
                
        # GUI input types
        self.inputFile = QLineEdit(self)
        self.outputFile = QLineEdit(self)
        self.malernummer = QLineEdit(self)
        self.P_Tema_kode = QComboBox(self)
        self.method_code = QComboBox(self)
        self.noyaktighet = QComboBox(self)
        self.synbarhet = QComboBox(self)
        self.auto_trase = QComboBox(self)

        # Input file search button
        inputButton = QPushButton("...")
        inputButton.clicked.connect(self.selectInfile)      
        inputButton.setToolTip('Select input file')

        # Output file search button
        outputButton = QPushButton("...")
        outputButton.clicked.connect(self.selectOutfile)
        outputButton.setToolTip('Select output folder')


        # OK button
        okButton = QPushButton("Confirm")
        okButton.clicked.connect(self.get_outfile_name)
        okButton.clicked.connect(self.read_file)
        okButton.setToolTip('Confirm settings')


        # Run button
        runButton = QPushButton("Convert!")
        runButton.clicked.connect(self.save_shpfiles)
        runButton.setToolTip('Save to SHP file')
        
       # Set layout as grid
        self.mainLayout = QGridLayout(self)

        # Define inputs
        self.mainLayout.addWidget(in_name_lab, 0, 0)
        self.mainLayout.addWidget(self.inputFile, 0, 1)
        self.inputFile.setPlaceholderText('Input.csv')
        self.inputFile.textChanged.connect(self.test_in)
        self.inputFile.textChanged.connect(self.auto_out)
        self.mainLayout.addWidget(inputButton, 0, 2)
        
        # Output 
        self.mainLayout.addWidget(out_name_lab, 1, 0)
        self.mainLayout.addWidget(self.outputFile, 1, 1)
        self.outputFile.setPlaceholderText('path/to/folder')
        self.mainLayout.addWidget(outputButton, 1, 2)
        
        # Username code 
        self.mainLayout.addWidget(num_lab, 2, 0)
        self.mainLayout.addWidget(self.malernummer)
        self.malernummer.setToolTip('Skriv inn måler og prosjekt nummer')
        self.malernummer.setPlaceholderText('MMnnnnnn')
        self.malernummer.setToolTip(
            '''Select the folder to save your shapefile into. 
            Default is the same folder in which your csv file is.'''
            )

        # PTEMA codes 
        self.mainLayout.addWidget(ptema_lab, 3, 0)
        self.mainLayout.addWidget(self.P_Tema_kode)
        self.ptema = [
            ('200 - Trase linje prosjektering', 200, 'Trase linje prosjektering'),
            ('201 - Trase nettstasjon prosjektering', 201, 'Trase nettstasjon prosjektering'),
            ('202 - Trase diverse prosjekteringprosk', 202, 'Trase diverse prosjekteringprosk'),
            ('210 - Trasepunkt mast prosjektert', 210, 'Trasepunkt mast prosjektert'),
            ('211 - Trasepunkt nettstasjon prosjektert', 211, 'Trasepunkt nettstasjon prosjektert'),
            ('212 - Trasepunkt diverse prosjektering', 212, 'Trasepunkt diverse prosjektering'),
            ('301 - Trase HSP 11-22kV', 301, 'Trase HSP 11-22kV'),
            ('303 - Trase veilys', 303, 'Trase veilys'),
            ('304 - Trase signal', 304, 'Trase signal'),
            ('305 - Trase kabelkjede', 305, 'Trase kabelkjede'),
            ('307 - Trase HSP 45-300kV', 307, 'Trase HSP 45-300kV'),
            ('312 - Trase LSP 240-400V', 312, 'Trase LSP 240-400V'),
            ('314 - Trase tomme rør', 314, 'Trase tomme rør'),
            ('316 - Trase ute av drift', 316, 'Trase ute av drift'),
            ('319 - Trasepunkt kabelkveil', 319, 'Trasepunkt kabelkveil'),
            ('320 - Trasepunkt kabelende', 320, 'Trasepunkt kabelende'),
            ('321 - Trase nettstasjon omriss', 321, 'Trase nettstasjon omriss'),
            ('324 - Trasepunkt landmålte punkt', 324, 'Trasepunkt landmålte punkt'),
            ('325 - Trase innm. situasjon', 325, 'Trase innm. situasjon'),
            ('326 - Trase rør med kabler', 326, 'Trase rør med kabler'),
            ('327 - Trasepunkt stålmast veilys', 327, 'Trasepunkt stålmast veilys'),
            ('329 - Trase avlastning plate', 329, 'Trase avlastning plate'),
            ('340 - Trase luft 45-420kV', 340, 'Trase luft 45-420kV'),
            ('341 - Trase luft 1-22kV', 341, 'Trase luft 1-22kV'),
            ('342 - Trase luft 0-1kV', 342, 'Trase luft 0-1kV'),
            ('350 - Trase luft ukjent', 350, 'Trase luft ukjent'),
            ('360 - Trasepunkt mast LSP', 360, 'Trasepunkt mast LSP'),
            ('380 - Trasepunkt koblingspunkt', 380, 'Trasepunkt koblingspunkt'),
            ('381 - Trasepunkt Skjøtende kveil', 381, 'Trasepunkt Skjøtende kveil'),
            ('399 - Trasepunkt fjernet', 399, 'Trasepunkt fjernet'),
            ('700 - Trasepunkt skjøt', 700, 'Trasepunkt skjøt'),
            ('701 - Trasepunkt kabelskap', 701, 'Trasepunkt kabelskap'),
            ('702 - Trasepunkt rørende', 702, 'Trasepunkt rørende'),
            ('703 - Trase borehull', 703, 'Trase borehull'),
            ('704 - Trasepunkt veilysskap', 704, 'Trasepunkt veilysskap'),
            ('705 - Trasepunkt signalskap', 705, 'Trasepunkt signalskap'),
            ('708 - Trase trekkekum omriss', 708, 'Trase trekkekum omriss'),
            ('712 - Trasepunkt kum kjede', 712, 'Trasepunkt kum kjede'),
            ('715 - Trase kjede ut av drift', 715, 'Trase kjede ut av drift'),
            ('716 - Trase kabel i kanal', 716, 'Trase kabel i kanal'),
            ('717 - Trase jordledning', 717, 'Trase jordledning'),
            ('720 - Trasepunkt bilder', 720, 'Trasepunkt bilder'),
            ('721 - Trasepunkt video', 721, 'Trasepunkt video'),
            ('777 - Trasepunkt innmålt TKS skap', 777, 'Trasepunkt innmålt TKS skap'),
            ('780 - Trase LSP usikker', 780, 'Trase LSP usikker'),
            ('781 - Trase veilys usikker', 781, 'Trase veilys usikker'),
            ('782 - Trase signal usikker', 782, 'Trase signal usikker'),
            ('783 - Trase rør usikker', 783, 'Trase rør usikker'),
            ('784 - Trasepunkt fordelingsskap usikker', 784, 'Trasepunkt fordelingsskap usikker'),
            ('785 - Trasepunkt veilysskap usikker', 785, 'Trasepunkt veilysskap usikker'),
            ('786 - Trasepunkt signalskap usikker', 786, 'Trasepunkt signalskap usikker'),
            ('787 - Trasepunkt mast LSP usikker', 787, 'Trasepunkt mast LSP usikker'),
            ('788 - Trasepunkt mast veilys usikker', 788, 'Trasepunkt mast veilys usikker'),
            ('789 - Trasepunkt skjøt usikker', 789, 'Trasepunkt skjøt usikker'),
            ('790 - Trasepunkt kabelkveil usikker', 790, 'Trasepunkt kabelkveil usikker'),
            ('791 - Trasepunkt kabelende usikker', 791, 'Trasepunkt kabelende usikker'),
            ('792 - Trase borehull usikker', 792, 'Trase borehull usikker'),
            ('793 - Trase ute av drift usikker', 793, 'Trase ute av drift usikker'),
            ('794 - Trase HSP usikker', 794, 'Trase HSP usikker'),
            ('796 - Trase tomme rør usikker', 796, 'Trase tomme rør usikker'),
            ('797 - Trasepunkt mast HSP', 797, 'Trasepunkt mast HSP'),
            ('798 - Trase nettstasjon usikker', 798, 'Trase nettstasjon usikker'),
            ('799 - Trase fjernet', 799, 'Trase fjernet'),
            ('900 - Trasepunkt Senterpunkt mast HSP', 900, 'Trasepunkt Senterpunkt mast HSP'),
            ('901 - Trasepunkt Mastebein HSP', 901, 'Trasepunkt Mastebein HSP'),
            ('902 - Trasepunkt Strever mast HSP', 902, 'Trasepunkt Strever mast HSP'),
            ('903 - Trasepunkt Bardunfestepunkt HSP', 903, 'Trasepunkt Bardunfestepunkt HSP'),
            ('904 - Trasepunkt Senterpunkt mast LSP', 904, 'Trasepunkt Senterpunkt mast LSP'),
            ('905 - Trasepunkt Mastebein LSP', 905, 'Trasepunkt Mastebein LSP'),
            ('906 - Trasepunkt Strever mast LSP', 906, 'Trasepunkt Strever mast LSP'),
            ('907 - Trasepunkt Bardunfestepunkt LSP', 907, 'Trasepunkt Bardunfestepunkt LSP'),
            ('908 - Trasepunkt Punkt ved maks pilhøyde', 908, 'Trasepunkt Punkt ved maks pilhøyde'),
            ('909 - Trasepunkt Annet punkt på linje', 909, 'Trasepunkt Annet punkt på linje'),
            ('910 - Trasepunkt Flymarkør', 910, 'Trasepunkt Flymarkør')
            ]
       
        self.LCODES = {
                        200,
                        201,
                        202,
                        301,
                        303,
                        304,
                        305,
                        307,
                        312,
                        314,
                        316,
                        321,
                        325,
                        326,
                        329,
                        340,
                        341,
                        342,
                        350,
                        703,
                        708,
                        715,
                        716,
                        717,
                        780,
                        781,
                        782,
                        783,
                        792,
                        793,
                        794,
                        796,
                        798,
                        799,
                        701, # - TERMINAL POINT CODES
                        784, #
                        360, #
                        787, #
                        321, #
                        798, #
                        797  #
                        }
        
        self.TERMINAL_CODES = {
                        701, # - TERMINAL POINT CODES
                        784, #
                        360, #
                        787, #
                        321, #
                        798, #
                        797  #
                        }
        
        self.ptema_list = [list[0] for list in self.ptema]
        self.P_Tema_kode.addItems(self.ptema_list)
        self.P_Tema_kode.setCurrentIndex(
            self.ptema_list.index('324 - Trasepunkt landmålte punkt')
            )
        self.P_Tema_kode.setToolTip('Select PTema')
        

        # Maalemethode codes
        self.mainLayout.addWidget(meth_lab, 4, 0)
        self.mainLayout.addWidget(self.method_code)
        self.method = [
            ('10 - Uspesifisert målemetode', 10, 'Uspesifisert målemetode'),
            ('11 - Totalstasjon', 11, 'Totalstasjon'),
            ('15 - Utmål fra bygglinje', 15, 'Utmål fra bygglinje'),
            ('47 - Digitalisert på skjem fra grunnkart FKB', 47, 'Digitalisert på skjem fra grunnkart FKB'),
            ('49 - Laserdata NN2000', 49, 'Laserdata NN2000'),
            ('51 - Digitalisert god skisse', 51, 'Digitalisert god skisse'),
            ('56 - Digitalisert fra scannet kart', 56, 'Digitalisert fra scannet kart'),
            ('82 - Digitalisert frihåndstegning', 82, 'Digitalisert frihåndstegning'),
            ('93 - GPS-Statisk', 93, 'GPS-Statisk'),
            ('96 - GPS-Rtk', 96, 'GPS-Rtk'),
            ('99 - Ukjent målemetode', 99, 'Ukjent målemetode'),
            ('966 - GPS-Rtk NC', 966, 'GPS-Rtk NC')
            ]
        self.method_list = [list[0] for list in self.method]
        self.method_code.addItems(self.method_list)
        self.method_code.setCurrentIndex(self.method_list.index('96 - GPS-Rtk'))
        self.method_code.setToolTip('Select målemetode')
        
        # Noyaktighet discreet values in cm
        self.mainLayout.addWidget(noyakt_lab, 5, 0)
        self.mainLayout.addWidget(self.noyaktighet)
        self.noy = [
            ('Auto', 9999),
            ('5', 5), 
            ('10', 10),
            ('20', 20), 
            ('30', 30), 
            ('100', 100), 
            ('200', 200), 
            ('300', 300), 
            ('400', 400), 
            ('500', 500), 
            ('1000', 1000)
            ]
        self.noy_list = [list[0] for list in self.noy]
        self.noyaktighet.addItems(self.noy_list)
        self.noyaktighet.setCurrentIndex(self.noy_list.index('Auto'))
        self.noyaktighet.setToolTip('Select noyaktighet')

        # Synbarhet codes
        self.mainLayout.addWidget(syn_lab, 6, 0)
        self.mainLayout.addWidget(self.synbarhet)
        self.syn = [
            ('0 - Fullt synlig ved innmåling', 0, 'Fullt synlig ved innmåling'), 
            ('1 - Innmåling lukket grøft', 1, 'Innmåling lukket grøft'), 
            ('3 - Ikke synlig trase; i sjø/undergrunn', 3, 'Ikke synlig trase; i sjø/undergrunn')
            ]
        self.syn_list = [list[0] for list in self.syn]
        self.synbarhet.addItems(self.syn_list)
        self.synbarhet.setCurrentIndex(self.syn_list.index('1 - Innmåling lukket grøft'))        
        self.synbarhet.setToolTip('Select Synbarhet')

        # Trase codes
        self.mainLayout.addWidget(trase_lab, 7, 0)
        self.mainLayout.addWidget(self.auto_trase)
        self.auto_trase.addItems(['Yes', 'No'])
        self.auto_trase.setCurrentIndex(1)        
        self.auto_trase.setToolTip('Automatic Trase')

        self.auto_trase.currentTextChanged.connect(self.toggle_auto_trase_options)

        # Auto-trase extra parameters (initially hidden)

        self.alg_label = QLabel("Algorithm:")
        self.alg_input = QComboBox(self)
        self.nn_alg = ['greedy', 'prim', 'kruskal', 'boruvka']
        self.alg_input.addItems(self.nn_alg)

        self.dist_scale_label = QLabel("Distance weighting:")
        self.dist_scale_input = QLineEdit("100")
        self.dist_scale_input.setToolTip('Scale to increase penalty for higher distances, i.e. makes closer points more desirable. (1 to inf)')

        self.angle_scale_label = QLabel("Angle weighting:")
        self.angle_scale_input = QLineEdit("10")
        self.angle_scale_input.setToolTip('Scale to increase penalty for larger angles, i.e. makes points in-line more desirable. (1 to inf)')

        self.max_dist_label = QLabel("Maximum distance (m):")
        self.max_dist_input = QLineEdit("10")
        self.max_dist_input.setToolTip('No line will be drawn between points above this distance. (0 to inf)')

        self.search_angle_label = QLabel("Search angle (°):")
        self.search_angle_input = QLineEdit("80")
        self.search_angle_input.setToolTip('No line will be drawn to a point outside this angle from the previous 2 points. (90 to 0)')

        self.usikkert_dist_label = QLabel("Usikkert distance (m):")
        self.usikkert_dist_input = QLineEdit("8")
        self.usikkert_dist_input.setToolTip('Points further than this distance apart will be classified as usikkert. (0 to inf)')

        self.noyaktighet_max_label = QLabel("Maximum nøyaktighet (cm):")
        self.noyaktighet_max_input = QLineEdit("100")
        self.noyaktighet_max_input.setToolTip('Points with a nøyaktighet above this value will be classified as usikkert. (0 to inf)')

        # Add to layout but initially hide
        row = 8  # After existing layout rows
        self.extra_widgets = [
            (self.alg_label, self.alg_input),
            (self.dist_scale_label, self.dist_scale_input),
            (self.angle_scale_label, self.angle_scale_input),
            (self.max_dist_label, self.max_dist_input),
            (self.search_angle_label, self.search_angle_input),
            (self.usikkert_dist_label, self.usikkert_dist_input),
            (self.noyaktighet_max_label, self.noyaktighet_max_input)
        ]

        for i, (label, input) in enumerate(self.extra_widgets):
            self.mainLayout.addWidget(label, row + i, 0)
            self.mainLayout.addWidget(input, row + i, 1)
            label.hide()
            input.hide()


        self.mainLayout.addWidget(okButton, 15, 1)
        self.mainLayout.addWidget(runButton, 15, 2)  

        # Matplotlib canvas for plotting
        self.figure = Figure(figsize=(100, 100))
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setMinimumWidth(700)  
        self.ax = self.figure.add_subplot(111)
        self.ax.set_aspect('equal')

        # Navigation toolbar
        self.nav_toolbar = NavigationToolbar(self.canvas, self)
        self.canvas.setVisible(False)
        self.nav_toolbar.setVisible(False)

        # Add to layout (far right)
        plot_col = 4
        self.mainLayout.addWidget(self.nav_toolbar, 0, plot_col, 1, 1)
        self.mainLayout.addWidget(self.canvas, 1, plot_col, 14, 1)


    def outfile_warn(self):
        QMessageBox.warning(
            self, 
            'Output Warning',
            "The selected output path does not exist, the output will be written to the input folder as a default.",
            buttons=QMessageBox.StandardButton.Ok| QMessageBox.StandardButton.Cancel,
            defaultButton=QMessageBox.StandardButton.Cancel
        )
    
    def name_warn(self):    
        QMessageBox.warning(
            self, 
            'Målernummer Warning',
            "The målernummer does not follow the regular format e.g. MÅ123456. Remember that this is required to be unique for correct import into ArcGIS...",
            buttons=QMessageBox.StandardButton.Ok
        )

    def noyaktighet_error(self):    
        QMessageBox.warning(
            self, 
            'Noyaktighet error',
            "The file you are converting is a KOF file. The noyaktighet needs to be defined.",
            buttons=QMessageBox.StandardButton.Close
        )

    def infile_error(self):    
        QMessageBox.critical(
            self, 
            'Path error',
            "Input file not found. Please check path/to/file.csv",
            buttons=QMessageBox.StandardButton.Close
        )

    def csv_error(self):    
        QMessageBox.critical(
            self, 
            'Input file error',
            '''There was a problem with the input file, 
            likely due to format of ".kof" file, 
            or differences of ".csv" file to VLoc3 version 
            "2025 - VMMAP Web - 2.20.2" \n\n\n {}'''.format(format_exc()),
            buttons=QMessageBox.StandardButton.Close
        )

    def trase_error(self):    
        QMessageBox.critical(
            self, 
            'Auto trase error',
            '''There was a problem with auto trase,
            this is likely due to incorrect PTEMA codes from the VLoc3. 
            This is also an experimental feature and may require editing after upload." \n\n\n {}'''.format(format_exc()),
            buttons=QMessageBox.StandardButton.Close
        )

    def complete(self):
        QMessageBox.information(
        self,
        "Shapefile Written",
        "Your shapefile has been written to {:}.".format(self.outputFile.text()),
        buttons=QMessageBox.StandardButton.Ok
        )     

    def tema(self):
        self.use_tema = QMessageBox.question(
            self, 
            "PTEMA Already in file", 
            "The input CSV seems to already have a PTEMA column. Do you wish to use these values instead?", 
            buttons=QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
            defaultButton=QMessageBox.StandardButton.Yes
            )

    def selectInfile(self):
        inputFile = QFileDialog.getOpenFileName(
            self, 
            "Finne målefila", 
            '', 
            "CSV files (*.csv *.kof *.xlsx *.KOF)"
            )
        inputPath = Path(inputFile[0])
        self.input_stem = str(inputPath.stem)
        self.input_string = str(inputPath)
        self.inputFile.setText(self.input_string)
        
        if Path(self.inputFile.text()).suffix in'.kof':
            self.noyaktighet.setCurrentIndex(self.noy_list.index('5'))
        else: 
            self.noyaktighet.setCurrentIndex(self.noy_list.index('Auto'))

    def selectOutfile(self):
        outputFolder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        outputPath = Path(outputFolder)
        self.outputFile.setText(str(outputPath))
    
    def test_in(self):
        if not os.path.isfile(self.inputFile.text()):
            self.infile_error()

    
    def get_outfile_name(self):      
        pattern = r'^[A-Za-z]{2}\d{6}$'
        if re.match(pattern, self.malernummer.text()):  # Check if malernummer is of valid inpuut pattern (XX######)
           
            if (self.outputFile.text().endswith(".shp") and 
            os.path.isdir(os.path.dirname(self.outputFile.text()))):    # If output is named already and the directory exists
                self.outfile_path = os.path.dirname(self.outputFile.text())   # Then just use that name and path

            elif os.path.isdir(self.outputFile.text()):     # Else if the path is just a directory and not a named file 
                self.outfile_path = Path(self.outputFile.text())
                self.outfile_name = self.input_stem + '_' + self.malernummer.text()     # Append the malernummer into the filename
                self.out_path = os.path.join(self.outfile_path, self.outfile_name)
                self.out_path = Path(self.out_path).with_suffix('.shp')     # Stick a suffix on it
                self.outputFile.setText(str(self.out_path))     # Change the GUI text
            
            elif not os.path.isdir(os.path.dirname(self.outputFile.text())):        
                # Else if the path is not a dir stick a flag up and set the file name back to default
                self.outfile_warn()
                self.auto_out()
        
        else:
            self.name_warn()        
    
    def auto_out(self): # Make the output name automatic
        self.outputFile.setText(os.path.dirname(self.inputFile.text()))
    

    def read_file(self):
        try:
            if Path(self.inputFile.text()).suffix in'.kof': 
                self.convert_kof()
            elif Path(self.inputFile.text()).suffix in '.csv':
                self.convert_csv()
        except Exception:
            self.csv_error()
            print(format_exc())

        if self.auto_trase.currentIndex() == 0:
            try:
                self.write_polylines()
                self.plot_trase(self.lines_gdf, self.shp_gdf)
            except Exception:
                self.trase_error()
                print(format_exc())

    
    def toggle_auto_trase_options(self, text):
        show = text == 'Yes'
        for label, input in self.extra_widgets:
            label.setVisible(show)
            input.setVisible(show)
        self.canvas.setVisible(show)
        self.nav_toolbar.setVisible(show)

    
    def plot_trase(self, lines_gdf, shp_gdf):
        self.ax.clear()

        lines_gdf = lines_gdf.to_crs(epsg=3857)
        shp_gdf = shp_gdf.to_crs(epsg=3857)

        for _, row in lines_gdf.iterrows():
            coords_2d = [(p[0], p[1]) for p in row.geometry.coords]
            xs, ys = zip(*coords_2d)

            if int(row['LTEMA']) in self.LSP_CABLE:
                color = 'green' 
            elif int(row['LTEMA']) in self.HSP_CABLE:
                color = 'red'

            linestyle = '--' if 'GOOD QUALITY' not in row['TEMA_REASON'] else '-'

            self.ax.plot(xs, ys, color=color, linestyle=linestyle, linewidth=2)

        # Draw terminal and non-terminal points
        for _, row in shp_gdf.iterrows():
            x, y = row.geometry.x, row.geometry.y
            code = int(row['LTEMA'])
            marker = 'o' if code in self.CABLE_CODES else 's'
            self.ax.plot(x, y, marker, color='black', markersize=4, markerfacecolor='black')
            self.ax.annotate(str(row['PNUMMER']), xy=(x,y))

        self.ax.set_title("Auto-Trase Preview")
        self.ax.set_xlabel("Easting")
        self.ax.set_ylabel("Northing")
        self.ax.tick_params(labelsize=8)

        # Hide tick labels if values aren't useful
        self.ax.set_xticklabels([])
        self.ax.set_yticklabels([])
        self.ax.axis('equal') 


        ctx.add_basemap(self.ax, source=ctx.providers.OpenStreetMap.Mapnik)  # Or use .Esri.WorldImagery for satellite

        self.canvas.draw()

    def save_shpfiles(self):
        shapefile_path = Path(self.outputFile.text())
        csv_path = shapefile_path.with_suffix(".csv")

        # Check if either file exists
        if shapefile_path.exists() or csv_path.exists():
            reply = QMessageBox.question(
                self,
                "Overwrite Confirmation",
                f"The file {shapefile_path.name} or its CSV already exists.\nDo you want to overwrite it?",
                buttons=QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                defaultButton=QMessageBox.StandardButton.No
            )

            if reply != QMessageBox.StandardButton.Yes:
                return  # Cancel operation

        for col in self.shp_gdf.columns:
            if col != self.shp_gdf.geometry.name:  # Skip the geometry column
                self.shp_gdf[col] = self.shp_gdf[col].astype(str)
        
        self.point_schema = gpd.io.file.infer_schema(self.shp_gdf)

        for field in ['PTEMA', 'LTEMA', 'MAKS_AVVIK']: # Set data type LONG in ARCMap
            self.point_schema['properties'][field] = 'int32:10'

        for field in ['MALEMETODE', 'SYNBARHET', 'H_MALEMETODE']: # Set data type SHORT in ARCMap
            self.point_schema['properties'][field] = 'int32:4'

        # Proceed with writing the files
        self.shp_gdf.to_file(shapefile_path, driver="ESRI Shapefile", schema=self.point_schema, engine="fiona")
        self.shp_gdf.to_csv(csv_path, index=False)
        
        if self.auto_trase.currentIndex() == 0:
            try:
                self.save_polylines()
            except Exception:
                self.trase_error()
                print(format_exc())

        self.complete() # Send 'completed' message

    def save_polylines(self):
        self.line_schema = gpd.io.file.infer_schema(self.lines_gdf)    
        for field in ['LTEMA']: # Set data type LONG in ARCMap
            self.line_schema['properties'][field] = 'int32:10'

        for field in ['MALEMETODE', 'SYNBARHET']: # Set data type SHORT in ARCMap
            self.line_schema['properties'][field] = 'int32:4'

        line_path = str(Path(self.outputFile.text()).with_suffix('')) + '_TRASE'
        line_path = Path(line_path).with_suffix('.shp')
        csv_path = Path(line_path).with_suffix('.csv')
        self.lines_gdf.to_file(line_path, driver="ESRI Shapefile", schema=self.line_schema, engine="fiona")
        self.lines_gdf.to_csv(csv_path, index=False)
                
    def convert_kof(self):
        def parse_kof_file(file_path):
            metadata_dict = {}
            points = []
            job_create_date = None

            with open(file_path, 'r', encoding='utf-8') as file:
                for raw_line in file:
                    line = raw_line.rstrip()

                    # Remove leading spaces before checking type
                    stripped = line.lstrip()

                    # Handle metadata
                    if stripped.startswith("00"):
                        content = stripped[2:].strip()
                        if ":" in content:
                            key, value = content.split(":", 1)
                            metadata_dict[key.strip()] = value.strip()
                            if key.strip() == "Job create Date":
                                job_create_date = value.strip()

                    # Handle points
                    elif stripped.startswith("05"):
                        # Pad to ensure line is long enough
                        line = line.ljust(60)
                        try:
                            pnummer = line[4:15].strip()
                            ptema = line[15:24].strip() or 324
                            for item in self.ptema:
                                if item[1] == int(ptema):
                                    tekst = item[2]
                            northing_str = line[24:37].strip()
                            easting_str = line[37:49].strip()
                            elevation_str = line[49:58].strip()

                            easting = float(easting_str) if easting_str else None
                            northing = float(northing_str) if northing_str else None
                            elevation = float(elevation_str) if elevation_str else None



                            points.append([pnummer, ptema, tekst, northing, easting, elevation])

                        except Exception as e:
                            print(f"Skipping malformed line: {line}\nReason: {e}")
                            continue
                        
                #print('POINTS {:}'.format(points))

            # Check if points were actually collected
            if not points:
                raise ValueError("No points were parsed. Check file formatting or line parsing logic.")

            # Create DataFrame
            points_df = pd.DataFrame(points, columns=["PNUMMER", "PTEMA", "TEMATEKST", "northing", "easting", "KOORDH"])

            # Add metadata fields
            for key, value in metadata_dict.items():
                points_df[key.replace(" ", "_")] = value

            # Add Job Create Date explicitly
            points_df["DATO"] = job_create_date

            # Add user-input fields
            points_df["LANDMALER"] = self.malernummer.text().upper()[:5]
            points_df["NOYAKTIGHE"] = self.noy[self.noyaktighet.currentIndex()][0]
            points_df["H_NOYAKTIGHE"] = self.noy[self.noyaktighet.currentIndex()][0]
            points_df["SYNBARHET"] = self.syn[self.synbarhet.currentIndex()][1]
            points_df["H_MALEMETODE"] = self.method[self.method_code.currentIndex()][1]
            points_df["MALEMETODE"] = self.method[self.method_code.currentIndex()][1]

            return metadata_dict, points_df

        # Example usage
        metadata, kof_df = parse_kof_file(self.inputFile.text())
        
        geometry = [
        Point(xyz) for xyz in zip(
        kof_df['easting'], 
        kof_df['northing'], 
        kof_df['KOORDH']
        )
        ]  

        self.shp_gdf = gpd.GeoDataFrame(kof_df, geometry=geometry, crs="EPSG:25832")     # GDF GO! 
        
        # Create a mask before modifying any columns
        mask = self.shp_gdf["PTEMA"].astype(int).isin(self.LCODES)

        # Create the new columns based on mask
        self.shp_gdf["LTEMA"] = np.where(mask, self.shp_gdf["PTEMA"], -1)
        self.shp_gdf["LTEMATEKST"] = np.where(mask, self.shp_gdf["TEMATEKST"], "")

        ptema_mask = (self.shp_gdf["PTEMA"].astype(int).isin(self.LCODES)) & (~self.shp_gdf["PTEMA"].astype(int).isin(self.TERMINAL_CODES))
        # Now update the original columns where PTEMA is in LCODES
        self.shp_gdf.loc[ptema_mask, "PTEMA"] = 324
        self.shp_gdf.loc[ptema_mask, "TEMATEKST"] = "Trasepunkt landmålte punkt"


        # Pythag to get maximum 3D error
        self.shp_gdf['MAKS_AVVIK'] =  np.ceil(
                                 np.sqrt(
                                      2*(
                                        self.shp_gdf['NOYAKTIGHE'].astype(int)**2
                                        )
                                        )
                                        )      

    def convert_csv(self):
        '''
        Take csv and covert it to a pandas dataframe,
        adjust for european vs english formattings. 

        Standardise column names, adjust names for versioning.

        Create gdf, make fields and do calculations.

        Write shapefile. 
        '''
        with open(Path(self.inputFile.text())) as file:
            dialect = csv.Sniffer().sniff(file.read(1024))
            delim = dialect.delimiter        
        
        # Define decimal based on delimiter for european formatting
        if delim == ';':
            dec = ','
        else:
            dec = '.'
        
        #print('\nDetected delimiter: {}\n'.format(delim))
        #print('\nDetected decimal: {}\n'.format(dec))

        try:
            csv_df = pd.read_csv(Path(self.inputFile.text()), sep = delim, decimal = dec)
        except: 
            csv_df = pd.read_excel(Path(self.inputFile.text()), sep = delim, decimal = dec)

        #print('\nORIGINAL COLUMNS:\n{}\n'.format(csv_df.columns))

        csv_df.columns = (
            csv_df.columns.str.strip().str.lower()
            .str.replace(" ", "")
            .str.replace("[()€$]", "", regex=True)
            )   # Make the headers as similar as possible for ease with versions 

        #print('\nNEW COLUMNS:\n{}'.format(csv_df.columns))


        for col in csv_df.columns:  # Swap the annoying minus symbol and try to make it numeric
            try:
                csv_df[col] = pd.to_numeric(csv_df[col].str.replace(',','.').str.replace('−','-'))   
            except:
                pass
            print('Column {} has dtype {}'.format(col, csv_df[col].dtype))

        # use the matching column name 
        northing_label = csv_df.columns[csv_df.columns.str.contains('northing')][0]
        easting_label  = csv_df.columns[csv_df.columns.str.contains('easting')][0]
        altitude_label = csv_df.columns[csv_df.columns.str.contains('altitude/')][0]
        depth_label = csv_df.columns[csv_df.columns.str.contains('depth')][0]
        time_label = csv_df.columns[csv_df.columns.str.contains('gpstime')][0]
        fix_label = csv_df.columns[csv_df.columns.str.contains('gpsfix')][0]
        bargraph_label = csv_df.columns[csv_df.columns.str.contains('bargraph%')][0]
        rms_label = csv_df.columns[csv_df.columns.str.contains('2drmsm')][0]
        offset_label = csv_df.columns[csv_df.columns.str.contains('offsetm')][0]
        vecsep_label = csv_df.columns[csv_df.columns.str.contains('vectorseparationm')][0]
       
        try:    # Look for features, for ptema allocation
            fcode_label = csv_df.columns[csv_df.columns.str.contains('featurecode')][0]
            fdesc_label = csv_df.columns[csv_df.columns.str.contains('featuredescription')][0]
        except:
            pass

        try:    # Formatting differences from old to new
            current_label = csv_df.columns[csv_df.columns.str.contains('locatecurrent')][0]
        except:
            current_label = csv_df.columns[csv_df.columns.str.contains('current')][0]
        
        try:    # Formatting differences from old to new
            gain_label = csv_df.columns[csv_df.columns.str.contains('locatorgaindb')][0]
        except:
            gain_label = csv_df.columns[csv_df.columns.str.contains('gaindb')][0]

        try:    # Use UP# point if available else make it with the record idex or then finally the index
            point_label = csv_df.columns[csv_df.columns.str.contains('up#')][0]
        except:
            try:
                point_label = csv_df.columns[csv_df.columns.str.contains('recordindex')][0]
                csv_df[point_label] = 'UP' + csv_df[point_label].astype(int).astype(str)
            except: 
                csv_df['up#'] = 'UP' + csv_df.index.astype(str)
                point_label = 'up#'        
        
        csv_df = csv_df.iloc[1:]    # Remove "Observations:" line
        csv_df = csv_df[csv_df[fix_label] != 'NONE']    # Clean points with no GPS fix
        csv_df = csv_df.copy()
        csv_df.reset_index(inplace=True)
        
        print(
            '''
            # ----------------------------------------------------------------------------\n
            # Drawing Shapefile\n
            # ----------------------------------------------------------------------------\n
            '''
            )

        csv_df['cablealtitude'] = csv_df[altitude_label] - csv_df[depth_label]  # Calculate the altitude of cable relative to sea-level

        geometry = [
            Point(xyz) for xyz in zip(
            csv_df[easting_label], 
            csv_df[northing_label], 
            csv_df['cablealtitude']
            )
            ]
        
        shp_gdf = gpd.GeoDataFrame(csv_df, geometry=geometry, crs="EPSG:25832")     # GDF GO!
        
        # Define the new fields in gdf
        shp_gdf['KOORDH'] = csv_df['cablealtitude']
        shp_gdf['DATO'] = csv_df[time_label]
        shp_gdf['PNUMMER'] = csv_df[point_label]
        shp_gdf['GPS_fix'] = csv_df[fix_label]
        shp_gdf['Current_mA'] = csv_df[current_label]
        shp_gdf['Gain_dB'] = csv_df[ gain_label]
        shp_gdf['Bargraph_%'] = csv_df[bargraph_label]
        shp_gdf['Depth'] = csv_df[depth_label]
        shp_gdf['2DRMS_m'] = csv_df[rms_label]
        shp_gdf['Offset_m'] = csv_df[offset_label]
        shp_gdf['Vec_sep_m']= csv_df[vecsep_label]

        shp_gdf['LANDMALER'] = self.malernummer.text().upper()[:5]
        shp_gdf['SYNBARHET'] = self.syn[self.synbarhet.currentIndex()][1]  
        shp_gdf['H_MALEMETODE'] = self.method[self.method_code.currentIndex()][1]
        shp_gdf['MALEMETODE'] = self.method[self.method_code.currentIndex()][1]

        
        if 'fcode_label' in locals():  # If the feature code field exists ask if it is to be used
            self.tema()
            if self.use_tema == QMessageBox.StandardButton.Yes: 
                try:
                    shp_gdf['PTEMA'] = csv_df[fcode_label]
                    shp_gdf['TEMATEKST'] = csv_df[fdesc_label]
                except:
                    shp_gdf['PTEMA'] = self.ptema[self.P_Tema_kode.currentIndex()][1]
                    shp_gdf['TEMATEKST'] = self.ptema[self.P_Tema_kode.currentIndex()][2]
            else:    
                shp_gdf['PTEMA'] = self.ptema[self.P_Tema_kode.currentIndex()][1]
                shp_gdf['TEMATEKST'] = self.ptema[self.P_Tema_kode.currentIndex()][2]
        else: 
            shp_gdf['PTEMA'] = self.ptema[self.P_Tema_kode.currentIndex()][1]
            shp_gdf['TEMATEKST'] = self.ptema[self.P_Tema_kode.currentIndex()][2]            

        if self.noy[self.noyaktighet.currentIndex()][0] == 'Auto':
            noy_values = [list[1] for list in self.noy[1::]]    # Get discreet values for noyaktighet 
            totalerror = np.sqrt(
                abs(csv_df[rms_label])**2 
                + abs(csv_df[offset_label])**2 
                + abs(csv_df[vecsep_label])**2
                ) * 100     # Root square error
            noyaktighet = [min(noy_values, key=lambda x: abs(x - val)) for val in totalerror]   # Take nearest discreet value
            csv_df['noyakt'] = [str(item) for item in noyaktighet]
        
        try:               
            shp_gdf['NOYAKTIGHE'] = csv_df['noyakt']
        except:
            shp_gdf['NOYAKTIGHE'] = self.noy[self.noyaktighet.currentIndex()][0]
        
        shp_gdf['H_NOYAKTIGHE'] = shp_gdf['NOYAKTIGHE']

        # Pythag to get maximum 3D error
        shp_gdf['MAKS_AVVIK'] =  np.ceil(
                                 np.sqrt(
                                      2*(
                                        shp_gdf['NOYAKTIGHE'].astype(int)**2
                                        )
                                        )
                                        )

        self.shp_gdf = shp_gdf[
            ['PNUMMER', 
            'DATO', 
            'LANDMALER', 
            'GPS_fix', 
            '2DRMS_m', 
            'Offset_m',
            'Vec_sep_m',
            'Current_mA', 
            'Gain_dB', 
            'Bargraph_%', 
            'Depth',
            'KOORDH', 
            'PTEMA', 
            'TEMATEKST',
            'MALEMETODE',
            'H_MALEMETODE', 
            'NOYAKTIGHE',
            'H_NOYAKTIGHE', 
            'SYNBARHET', 
            'MAKS_AVVIK',
            'geometry']
            ].copy()
        
        # Create a mask before modifying any columns
        mask = self.shp_gdf["PTEMA"].astype(int).isin(self.LCODES)

        # Create the new columns based on mask
        self.shp_gdf["LTEMA"] = np.where(mask, self.shp_gdf["PTEMA"], -1)
        self.shp_gdf["LTEMATEKST"] = np.where(mask, self.shp_gdf["TEMATEKST"], "")

        ptema_mask = (self.shp_gdf["PTEMA"].astype(int).isin(self.LCODES)) & (~self.shp_gdf["PTEMA"].astype(int).isin(self.TERMINAL_CODES))
        # Now update the original columns where PTEMA is in LCODES
        self.shp_gdf.loc[ptema_mask, "PTEMA"] = 324
        self.shp_gdf.loc[ptema_mask, "TEMATEKST"] = "Trasepunkt landmålte punkt"

    def write_polylines(self):
        '''
        Function to automatically trace cable lines over points following common patterns in surveying process.

        VERY EXPERIMENTAL AND REQUIRES EDITING AFTER USE. 

        Uses a weighted distance by comperint the angle between the previous point lines and seeing which lay on a 'straighter' path.
        This the scales along with the 2D distance to weight a nearest neighbour algorithm.

        The algorithm also labels line segments as usikkert based on the noyaktighet of the poinst used to draw that line. 
        '''
        
        DIST_SCALE = int(self.dist_scale_input.text()) #distance scale
        ANGLE_SCALE = int(self.angle_scale_input.text()) #angle scale
        MAX_DIST = float(self.max_dist_input.text()) #m
        SEARCH_ANGLE = float(self.search_angle_input.text()) #deg
        USIKKERT_DIST = float(self.usikkert_dist_input.text()) #m
        NOYAKTIGHET_MAX = float(self.noyaktighet_max_input.text()) #cm

        LSP_240V = {312}
        LSP_USIKKER = {780}
        self.LSP_CABLE = LSP_240V | LSP_USIKKER
        LSP_MAST = {360, 787}
        SKAP_CODES = {701, 784}

        HSP_11KV = {301}
        HSP_45KV = {307}
        HSP_USIKKER = {794}
        HSP_MAST = {797}
        HSP_11KV_CABLE = HSP_11KV | HSP_USIKKER 
        HSP_45KV_CABLE = HSP_45KV | HSP_USIKKER 
        self.HSP_CABLE = HSP_11KV | HSP_45KV | HSP_USIKKER 
        
        NETTSTASJON_CODES = {321, 798}
        USIKKERT_PCODES = {780, 784, 794, 798}
        
        self.CABLE_CODES = self.LSP_CABLE | HSP_11KV | HSP_45KV | HSP_USIKKER
        
        LSP_NETWORK = self.LSP_CABLE | LSP_MAST | SKAP_CODES | NETTSTASJON_CODES
        LSP_TERMINALS = LSP_MAST | SKAP_CODES | NETTSTASJON_CODES

        HSP_NETWORK = self.HSP_CABLE | NETTSTASJON_CODES | HSP_MAST
        HSP_11KV_NETWORK = HSP_11KV_CABLE | NETTSTASJON_CODES | HSP_MAST
        HSP_45KV_NETWORK = HSP_45KV_CABLE | NETTSTASJON_CODES | HSP_MAST       
        HSP_TERMINALS = NETTSTASJON_CODES | HSP_MAST

        self.TERMINALS = LSP_TERMINALS | HSP_TERMINALS

        MASTS = LSP_MAST | HSP_MAST

        TERMINAL_DIST = 5
  
        def can_connect(ltema1, ltema2):
            # Is HSP11kv
            if (ltema1 in HSP_11KV_CABLE and ltema2 in HSP_11KV_NETWORK) or (ltema1 in HSP_11KV_NETWORK and ltema2 in HSP_11KV_CABLE):
                return True
            
            # Is HSP45kv
            if (ltema1 in HSP_45KV_CABLE and ltema2 in  HSP_45KV_NETWORK) or (ltema1 in  HSP_45KV_NETWORK and ltema2 in HSP_45KV_CABLE):
                return True
            
            # Is LSP
            if (ltema1 in self.LSP_CABLE and ltema2 in LSP_NETWORK) or (ltema1 in LSP_NETWORK and ltema2 in self.LSP_CABLE):
                return True
            

        # Extract 3D coordinates
        self.shp_gdf['x'] = self.shp_gdf.geometry.apply(lambda p: p.x)
        self.shp_gdf['y'] = self.shp_gdf.geometry.apply(lambda p: p.y)
        self.shp_gdf['z'] = self.shp_gdf.geometry.apply(lambda p: p.z)
        coords = self.shp_gdf[['x', 'y', 'z']].values

        terminal_indices = self.shp_gdf[self.shp_gdf['LTEMA'].isin(self.TERMINALS)].index
        terminal_coords = coords[terminal_indices]

        try:
            db = DBSCAN(eps=TERMINAL_DIST, min_samples=1).fit(terminal_coords)
            terminal_cluster_labels = db.labels_

            # Map: cluster_id -> list of global point indices in that cluster
            terminal_clusters = defaultdict(list)
            for idx, cluster_id in zip(terminal_indices, terminal_cluster_labels):
                terminal_clusters[cluster_id].append(idx)

            # For each cluster, retain only one point as the valid "connection target"
            valid_terminal_points = set()
            for cluster in terminal_clusters.values():
                # Pick representative (e.g., point with lowest index or most central)
                representative = min(cluster, key=lambda idx: np.mean(np.linalg.norm(coords[cluster] - coords[idx], axis=1)))
                valid_terminal_points.add(representative)
        except:
            valid_terminal_points = set()

        # Using sqrt reduces computing times and scales nicely when points are not abundant
        nbrs = NearestNeighbors(n_neighbors=int(np.sqrt(len(self.shp_gdf))) ,metric='euclidean').fit(coords) 
        distances, indices = nbrs.kneighbors(coords)
        G = nx.Graph()

        def compute_cosine_weight(p_prev, p_curr, p_next, angle_scale, dist_scale):
            """
            Draws possible lines to previous and potential next points and compares angle differences of lines
            ruling out sharp angles and giving favourable scaling to straighter paths.
            """
            v1 = p_curr - p_prev
            v2 = p_next - p_curr

            angle_threshold = SEARCH_ANGLE

            norm1 = np.linalg.norm(v1)
            norm2 = np.linalg.norm(v2)
            if norm1 == 0 or norm2 == 0:
                return np.inf, np.inf, angle_threshold  # fallback

            cos_theta = np.dot(v1, v2) / (norm1 * norm2)
            cos_theta = np.clip(cos_theta, -1, 1)
            theta = np.arccos(cos_theta)
            theta_deg = np.degrees(theta)

            if theta_deg >= angle_threshold:
                return np.inf, theta_deg, angle_threshold
            else:
                # Heavily penalize sharp angles
                return (1 / cos_theta ** angle_scale), theta_deg, angle_threshold
        
        edge_list = []
        for i, neighbors in enumerate(indices):
            for j in neighbors[1:]:
                # raw distance
                dist = np.linalg.norm(coords[i] - coords[j])
                # Optional: try to infer a previous point for direction (pick the closest one before i)
                prev_candidates = [k for k in neighbors[1:] if k != j and k != i]
                for prev in prev_candidates:
                    
                    dist2 = np.linalg.norm(coords[prev] - coords[i])
                    # compute direction-weighted distance
                    scale_fwd, angle_fwd, _ = compute_cosine_weight(coords[prev], coords[i], coords[j], ANGLE_SCALE, DIST_SCALE)
                    scale_rev, angle_rev, _ = compute_cosine_weight(coords[j], coords[i], coords[prev], ANGLE_SCALE, DIST_SCALE)
         
                    # If both pass, use the average scale for better symmetry
                    scale = (scale_fwd + scale_rev) / 2

                    weighted_dist = (dist**DIST_SCALE) * scale

                    '''print('Prev is {:}, I is {:}, J is {:}, dist is {:}, ' \
                    'scale is {:}, wdist is {:}, angle is {:}, ang.lim: {:}'
                    .format(prev,i,j,round(dist,3),round(scale,3),round(weighted_dist,3),round(angle,3),round(cutoff,3)))'''

                    ltema_i = int(self.shp_gdf['LTEMA'][i])
                    ltema_j = int(self.shp_gdf['LTEMA'][j]) 

                    if not can_connect(ltema_i, ltema_j):
                        print(f"Rejected: LTEMA {ltema_i} and {ltema_j} not connectable")
                    elif angle_fwd >= SEARCH_ANGLE or angle_rev >= SEARCH_ANGLE:
                        print(f"Rejected: angle too sharp — fwd {angle_fwd:.1f}, rev {angle_rev:.1f}")
                    elif dist > MAX_DIST or dist2 > MAX_DIST:
                        print(f"Rejected: distance too long — dist {dist:.1f}, dist2 {dist2:.1f}")
                    else:
                        print(f"Accepted: {i}-{j}, LTEMA {ltema_i}, dist {dist:.1f}, angle {angle_fwd:.1f}/{angle_rev:.1f}")
                    
                    if ltema_i in self.TERMINALS and i not in valid_terminal_points:
                        continue  # skip extra terminal points
                    if ltema_j in self.TERMINALS and j not in valid_terminal_points:
                        continue

                    if can_connect(ltema_i, ltema_j):
                        if angle_fwd < SEARCH_ANGLE and angle_rev < SEARCH_ANGLE:
                            if dist <= MAX_DIST and dist2 <= MAX_DIST:
                                edge_list.append([int(prev), int(i), int(j), float(weighted_dist)])
                                G.add_edge(i, j, weight=weighted_dist)
        
        if self.nn_alg[self.noyaktighet.currentIndex()] == 'greedy':
            # Build a greedy traversal graph: connect lowest-weight edges first
            sorted_edges = sorted(G.edges(data=True), key=lambda e: e[2]['weight'])
            graph = nx.Graph()
            uf = UnionFind()

            for u, v, data in sorted_edges:
                if uf[u] != uf[v]:  # only add if it doesn’t form a cycle
                    graph.add_edge(u, v, **data)
                    uf.union(u, v)
        else:
            graph = nx.minimum_spanning_tree(G, weight='weight', algorithm=self.nn_alg[self.noyaktighet.currentIndex()])


        # Tag each edge with its LTEMA
        edge_ltema = {}
        for u, v in graph.edges():
            ltema_u = int(self.shp_gdf['LTEMA'][u])
            ltema_v = int(self.shp_gdf['LTEMA'][v])
            noyakt_u = int(self.shp_gdf['NOYAKTIGHE'][u])
            noyakt_v = int(self.shp_gdf['NOYAKTIGHE'][v])
            raw_dist = np.linalg.norm(coords[u] - coords[v])

            is_usikkert = (
                ltema_u in USIKKERT_PCODES or 
                ltema_v in USIKKERT_PCODES or
                noyakt_u > NOYAKTIGHET_MAX or 
                noyakt_v > NOYAKTIGHET_MAX or 
                raw_dist > USIKKERT_DIST
            )
            
            reasons = []
            if is_usikkert:
                if ltema_u in USIKKERT_PCODES or ltema_v in USIKKERT_PCODES:
                    reasons.append("POINT DEFINED")
                if raw_dist > USIKKERT_DIST:
                    reasons.append(f"DIST > {USIKKERT_DIST}m")
                if noyakt_u > NOYAKTIGHET_MAX or noyakt_v > NOYAKTIGHET_MAX:
                    reasons.append(f"NOYAKT > {NOYAKTIGHET_MAX}cm")

            # Determine primary cable category
            cable_set = [ltema_u, ltema_v]
            if all(code in HSP_45KV_NETWORK for code in cable_set):
                ltema = 794 if is_usikkert else 307
            elif all(code in HSP_11KV_NETWORK for code in cable_set):
                ltema = 794 if is_usikkert else 301
            elif all(code in LSP_NETWORK for code in cable_set):
                ltema = 780 if is_usikkert else 312

            edge_ltema[(u, v)] = (ltema, reasons)
            edge_ltema[(v, u)] = (ltema, reasons)  # undirected


        # Traverse MST and merge segments with same LTEMA
        visited_edges = set()
        visited_nodes = set()

        merged_geoms = []
        merged_attrs = []

        def walk_linear_path(u, v, ltema):
            path = [u, v]
            collected_reasons = set(edge_ltema.get((u, v), (ltema, []))[1])
            visited_edges.add((u, v))
            visited_edges.add((v, u))

            # Walk in both directions
            def extend_path(start, prev):
                local_path = []
                curr = start
                while True:
                    curr_ltema = int(self.shp_gdf['LTEMA'][curr])
                    if curr_ltema in self.TERMINALS:
                        print(f"-= Breaking at STATION (LTEMA={curr_ltema}) at node {curr}")
                        break
                    if graph.degree[curr] != 2:
                        print(f"-< Breaking at JUNCTION (degree={graph.degree[curr]}) at node {curr}")
                        break

                    neighbors = [
                        n for n in graph.neighbors(curr)
                        if n != prev
                        and edge_ltema.get((curr, n), (None, []))[0] == ltema
                        and (curr, n) not in visited_edges
                    ]

                    if len(neighbors) != 1:
                        print(f"-| Breaking at FORK or END (found {len(neighbors)} next nodes) at node {curr}")
                        break

                    nxt = neighbors[0]
                    visited_edges.add((curr, nxt))
                    visited_edges.add((nxt, curr))
                    collected_reasons.update(edge_ltema.get((curr, nxt), (ltema, []))[1])
                    local_path.append(nxt)
                    prev, curr = curr, nxt
                return local_path

            left = extend_path(u, v)
            right = extend_path(v, u)

            full_path = list(reversed(left)) + path + right

            for node in path:
                visited_nodes.add(node)
                if int(self.shp_gdf['LTEMA'][node]) in MASTS:
                    break  # stop adding beyond this mast

            return full_path, collected_reasons

        for u, v in graph.edges():
            if (u, v) in visited_edges:
                continue
            ltema, _ = edge_ltema[(u, v)]
            path, reasons = walk_linear_path(u, v, ltema)

            line = LineString([Point(coords[i]) for i in path])
            point_data = self.shp_gdf.loc[path]

            max_noy = point_data['NOYAKTIGHE'].astype(int).max()
            malemetode = Counter(point_data['MALEMETODE']).most_common(1)[0][0]
            synbarhet = Counter(point_data['SYNBARHET']).most_common(1)[0][0]
            dato = point_data['DATO'].max()

            tematekst = "UNKNOWN"
            for item in self.ptema:
                if item[1] == ltema:
                    tematekst = item[2]
                    break
                    
            reason_str = ", ".join(sorted(reasons)) if reasons else "GOOD QUALITY"

            merged_geoms.append(line)
            merged_attrs.append({
                'LTEMA': str(ltema),
                'TEMATEKST': str(tematekst),
                'NOYAKTIGHE': str(max_noy),
                'MALEMETODE': str(malemetode),
                'SYNBARHET': str(synbarhet),
                'DATO': str(dato),
                'OBJTYPE': 'EL_Trase',
                'TEMA_REASON': str(reason_str)
            })

        # Replace old GeoDataFrame creation
        self.lines_gdf = gpd.GeoDataFrame(merged_attrs, geometry=merged_geoms, crs=self.shp_gdf.crs)

if __name__=="__main__":
    import sys
    app    = QApplication(sys.argv)
    myshow = InputDialog()
    myshow.setWindowTitle("CSV2SHP")
    app.setWindowIcon(QIcon('CSV2SHP.ico'))
    myshow.setWindowIcon(QIcon('CSV2SHP.ico'))
    myshow.show()
    app.exec()
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
from PyQt6.QtWidgets import QWidget, QPushButton, QLabel, QLineEdit, QComboBox, QFileDialog, QGridLayout,  QApplication, QErrorMessage, QMessageBox
from PyQt6.QtGui import QIcon
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
import numpy as np


# Base python packages
from traceback import format_exc
from pathlib import Path
from ctypes import windll
import os
import re
import csv


#### START ####
myappid = 'BKK.CSV2SHP.v1.3'
windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

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
                
        # GUI input types
        self.inputFile = QLineEdit(self)
        self.outputFile = QLineEdit(self)
        self.malernummer = QLineEdit(self)
        self.P_Tema_kode = QComboBox(self)
        self.method_code = QComboBox(self)
        self.noyaktighet = QComboBox(self)
        self.synbarhet = QComboBox(self)

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
        okButton.setToolTip('Confirm settings')


        # Run button
        runButton = QPushButton("Convert!")
        runButton.clicked.connect(self.read_file)
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
            ('210 - Trasepunkt mast prosjektert', 210, 'Trasepunkt mast prosjektert'),
            ('319 - Trasepunkt kabelkveil', 319, 'Trasepunkt kabelkveil'),
            ('320 - Trasepunkt kabelende', 320, 'Trasepunkt kabelende'),
            ('324 - Landmålte punkt', 324, 'Landmålte punkt'),
            ('327 - Trasepunkt stålmast veilys', 327, 'Trasepunkt stålmast veilys'),
            ('360 - Trasepunkt mast LSP', 360, 'Trasepunkt mast LSP'),
            ('399 - Trasepunkt fjernet', 399, 'Trasepunkt fjernet'),
            ('700 - Trasepunkt skjøt', 700, 'Trasepunkt skjøt'),
            ('701 - Trasepunkt kabelskap', 701, 'Trasepunkt kabelskap'),
            ('702 - Skal ikke brukes', 702, 'Skal ikke brukes'),
            ('704 - Trasepunkt veilysskap', 704, 'Trasepunkt veilysskap'),
            ('705 - Trasepunkt signalskap', 705, 'Trasepunkt signalskap'),
            ('712 - Trasepunkt kum kjede', 712, 'Trasepunkt kum kjede'),
            ('720 - Trasepunkt bilder', 720, 'Trasepunkt bilder'),
            ('721 - Video punkt', 721, 'Video punkt'),
            ('777 - Trasepunkt innmålt TKS skap', 777, 'Trasepunkt innmålt TKS skap'),
            ('784 - Trasepunkt fordelingsskap usikker', 784, 'Trasepunkt fordelingsskap usikker'),
            ('785 - Trasepunkt veilysskap usikker', 785, 'Trasepunkt veilysskap usikker'),
            ('786 - Trasepunkt signalskap usikker', 786, 'Trasepunkt signalskap usikker'),
            ('787 - Trasepunkt mast LSP usikker', 787, 'Trasepunkt mast LSP usikker'),
            ('788 - Trasepunkt mast veilys usikker', 788, 'Trasepunkt mast veilys usikker'),
            ('789 - Trasepunkt skjøt usikker', 789, 'Trasepunkt skjøt usikker'),
            ('790 - Trasepunkt kabelkveil usikker', 790, 'Trasepunkt kabelkveil usikker'),
            ('791 - Trasepunkt kabelende usikker', 791, 'Trasepunkt kabelende usikker'),
            ('797 - Trasepunkt mast HSP', 797, 'Trasepunkt mast HSP'),
            ('910 - Flymarkør', 910, 'Flymarkør'),
            ('900 - Senterpunkt mast HSP', 900, 'Senterpunkt mast HSP'),
            ('901 - Mastebein HSP', 901, 'Mastebein HSP'),
            ('902 - Strever mast HSP', 902, 'Strever mast HSP'),
            ('903 - Bardunfestepunkt HSP', 903, 'Bardunfestepunkt HSP'),
            ('904 - Senterpunkt mast LSP', 904, 'Senterpunkt mast LSP'),
            ('905 - Mastebein LSP', 905, 'Mastebein LSP'),
            ('906 - Strever mast LSP', 906, 'Strever mast LSP'),
            ('907 - Bardunfestepunkt LSP', 907, 'Bardunfestepunkt LSP'),
            ('908 - Annet', 908, 'Annet')
            ]
        self.ptema_list = [list[0] for list in self.ptema]
        self.P_Tema_kode.addItems(self.ptema_list)
        self.P_Tema_kode.setCurrentIndex(
            self.ptema_list.index('324 - Landmålte punkt')
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
            ('3 - Ikke synlig trase; i sjø/undergrunn', 2, 'Ikke synlig trase; i sjø/undergrunn')
            ]
        self.syn_list = [list[0] for list in self.syn]
        self.synbarhet.addItems(self.syn_list)
        self.synbarhet.setCurrentIndex(self.syn_list.index('1 - Innmåling lukket grøft'))        
        self.synbarhet.setToolTip('Select Synbarhet')


        self.mainLayout.addWidget(okButton, 7, 1)
        self.mainLayout.addWidget(runButton, 7, 2)

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
            "2025 - VMMAP Web - 2.19.6" \n\n\n {}'''.format(format_exc()),
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
            "Text files (*.csv) ;; Excel files (*.xlsx) ;; KOF files (*.kof)"
            )
        inputPath = Path(inputFile[0])
        self.input_stem = str(inputPath.stem)
        self.input_string = str(inputPath)
        self.inputFile.setText(self.input_string)

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
        
    
    def auto_out(self):
        self.outputFile.setText(os.path.dirname(self.inputFile.text()))
    

    def read_file(self):
        try:
            if Path(self.inputFile.text()).suffix in ['.kof', '.KOF']:
                self.convert_kof()
            elif Path(self.inputFile.text()).suffix in '.csv':
                self.convert_csv()
        except Exception:
            self.csv_error()

    def save_shpfiles(self, shp_gdf, schema):
        shapefile_path = Path(self.outputFile.text())
        csv_path = shapefile_path.with_suffix('.csv')

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

        # Proceed with writing the files
        shp_gdf.to_file(shapefile_path, driver="ESRI Shapefile", schema=schema, engine="fiona")
        shp_gdf.to_csv(csv_path, index=False)
        
        self.complete() # Send 'completed' message
                
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
        
        print('\nDetected delimiter: {}\n'.format(delim))
        print('\nDetected decimal: {}\n'.format(dec))

        try:
            csv_df = pd.read_csv(Path(self.inputFile.text()), sep = delim, decimal = dec)
        except: 
            csv_df = pd.read_excel(Path(self.inputFile.text()), sep = delim, decimal = dec)

        print('\nORIGINAL COLUMNS:\n{}\n'.format(csv_df.columns))

        csv_df.columns = (
            csv_df.columns.str.strip().str.lower()
            .str.replace(" ", "")
            .str.replace("[()€$]", "", regex=True)
            )   # Make the headers as similar as possible for ease with versions 

        print('\nNEW COLUMNS:\n{}'.format(csv_df.columns))


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
        
        print(self.shp_gdf.head())

        self.schema = gpd.io.file.infer_schema(self.shp_gdf)

        for field in ['PTEMA', 'MAKS_AVVIK']: # Set data type LONG in ARCMap
            self.schema['properties'][field] = 'int32:10'

        for field in ['MALEMETODE', 'SYNBARHET', 'H_MALEMETODE']: # Set data type SHORT in ARCMap
            self.schema['properties'][field] = 'int32:4'

        self.save_shpfiles(self.shp_gdf, self.schema)

        
        
if __name__=="__main__":
    import sys
    app    = QApplication(sys.argv)
    myshow = InputDialog()
    myshow.setWindowTitle("CSV2SHP")
    app.setWindowIcon(QIcon('CSV2SHP.ico'))
    myshow.setWindowIcon(QIcon('CSV2SHP.ico'))
    myshow.show()
    app.exec()
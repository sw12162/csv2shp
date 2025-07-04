#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on 25/06/2025 08:12

@author: sw12162
"""

from os import path


# Python script import
#import executables, conversions

# Externally installed packages
try:
    import pyi_splash
except:
    pass
from PyQt6.QtWidgets import QWidget, QPushButton, QLabel, QLineEdit, QComboBox, QFileDialog, QGridLayout,  QApplication, QErrorMessage, QMessageBox
from PyQt6.QtGui import QIcon
import geopandas as gpd
import pandas as pd
from shapely.geometry import Point

# Base python packages
from traceback import format_exc
from pathlib import Path
from ctypes import windll
import os
import re
from pathlib import Path
import csv

#### START ####

myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

try:
    pyi_splash.close()
except: 
    pass


class InputDialog(QWidget):
    def __init__(self):       
        super(InputDialog,self).__init__()
        
        #self.username = os.getlogin()

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
        

        self.mainLayout.addWidget(out_name_lab, 1, 0)
        self.mainLayout.addWidget(self.outputFile, 1, 1)
        self.outputFile.setPlaceholderText('path/to/folder')
        self.mainLayout.addWidget(outputButton, 1, 2)
        
        

        self.mainLayout.addWidget(num_lab, 2, 0)
        self.mainLayout.addWidget(self.malernummer)
        self.malernummer.setToolTip('Skriv inn måler og prosjekt nummer')
        self.malernummer.setPlaceholderText('MMnnnnnn')
        #self.malernummer.setText(self.username)
        self.malernummer.setToolTip('Select the folder to save your shapefile into. Default is the same folder in which your csv file is.')


        self.mainLayout.addWidget(ptema_lab, 3, 0)
        self.mainLayout.addWidget(self.P_Tema_kode)
        self.ptema = [('210 - Trasepunkt mast prosjektert', 210, '210 - Trasepunkt_mast_prosjektert', 'Trasepunkt mast prosjektert'), 
                      ('319 - Trasepunkt kabelkveil', 319, '319 - Trasepunkt_kabelkveil', 'Trasepunkt kabelkveil'), 
                      ('320 - Trasepunkt kabelende', 320, '320 - Trasepunkt_kabelende', 'Trasepunkt kabelende'), 
                      ('324 - Landmålte punkt', 324, '324 - Landmålte_punkt', 'Landmålte punkt'), 
                      ('327 - Trasepunkt stålmast veilys', 327, '327 - Trasepunkt_stålmast_veilys', 'Trasepunkt stålmast veilys'), 
                      ('360 - Trasepunkt mast LSP', 360, '360 - Trasepunkt_mast_LSP', 'Trasepunkt mast LSP'), 
                      ('399 - Trasepunkt fjernet', 399, '399 - Trasepunkt_fjernet', 'Trasepunkt fjernet'), 
                      ('700 - Trasepunkt skjøt', 700, '700 - Trasepunkt_skjøt', 'Trasepunkt skjøt'), 
                      ('701 - Trasepunkt kabelskap', 701, '701 - Trasepunkt_kabelskap', 'Trasepunkt kabelskap'), 
                      ('702 - Skal ikke brukes', 702, '702 - Skal_ikke_brukes', 'Skal ikke brukes'), 
                      ('704 - Trasepunkt veilysskap', 704, '704 - Trasepunkt_veilysskap', 'Trasepunkt veilysskap'), 
                      ('705 - Trasepunkt signalskap', 705, '705 - Trasepunkt_signalskap', 'Trasepunkt signalskap'), 
                      ('712 - Trasepunkt kum kjede', 712, '712 - Trasepunkt_kum_kjede', 'Trasepunkt kum kjede'), 
                      ('720 - Trasepunkt bilder', 720, '720 - Trasepunkt_bilder', 'Trasepunkt bilder'), 
                      ('721 - Video punkt', 721, '721 - Video_punkt', 'Video punkt'), 
                      ('777 - Trasepunkt innmålt TKS skap', 777, '777 - Trasepunkt_innmålt_TKS_skap', 'Trasepunkt innmålt TKS skap'), 
                      ('784 - Trasepunkt fordelingsskap usikker', 784, '784 - Trasepunkt_fordelingsskap_usikker', 'Trasepunkt fordelingsskap usikker'), 
                      ('785 - Trasepunkt veilysskap usikker', 785, '785 - Trasepunkt_veilysskap_usikker', 'Trasepunkt veilysskap usikker'), 
                      ('786 - Trasepunkt signalskap usikker', 786, '786 - Trasepunkt_signalskap_usikker', 'Trasepunkt signalskap usikker'), 
                      ('787 - Trasepunkt mast LSP usikker', 787, '787 - Trasepunkt_mast_LSP_usikker', 'Trasepunkt mast LSP usikker'), 
                      ('788 - Trasepunkt mast veilys usikker', 788, '788 - Trasepunkt_mast_veilys_usikker', 'Trasepunkt mast veilys usikker'), 
                      ('789 - Trasepunkt skjøt usikker', 789, '789 - Trasepunkt_skjøt_usikker', 'Trasepunkt skjøt usikker'), 
                      ('790 - Trasepunkt kabelkveil usikker', 790, '790 - Trasepunkt_kabelkveil_usikker', 'Trasepunkt kabelkveil usikker'), 
                      ('791 - Trasepunkt kabelende usikker', 791, '791 - Trasepunkt_kabelende_usikker', 'Trasepunkt kabelende usikker'), 
                      ('797 - Trasepunkt mast HSP', 797, '797 - Trasepunkt_mast_HSP', 'Trasepunkt mast HSP'), 
                      ('910 - Flymarkør', 910, '910 - Flymarkør', 'Flymarkør'), 
                      ('900 - Senterpunkt mast HSP', 900, '900 - Senterpunkt_mast_HSP', 'Senterpunkt mast HSP'), 
                      ('901 - Mastebein HSP', 901, '901 - Mastebein_HSP', 'Mastebein HSP'), 
                      ('902 - Strever mast HSP', 902, '902 - Strever_mast_HSP', 'Strever mast HSP'), 
                      ('903 - Bardunfestepunkt HSP', 903, '903 - Bardunfestepunkt_HSP', 'Bardunfestepunkt HSP'), 
                      ('904 - Senterpunkt mast LSP', 904, '904 - Senterpunkt_mast_LSP', 'Senterpunkt mast LSP'), 
                      ('905 - Mastebein LSP', 905, '905 - Mastebein_LSP', 'Mastebein LSP'), 
                      ('906 - Strever mast LSP', 906, '906 - Strever_mast_LSP', 'Strever mast LSP'), 
                      ('907 - Bardunfestepunkt LSP', 907, '907 - Bardunfestepunkt_LSP', 'Bardunfestepunkt LSP'), 
                      ('908 - Annet', 908, '908 - Annet', 'Annet')]
        
        self.P_Tema_kode.addItems([list[0] for list in self.ptema])
        self.ptema_list = [list[0] for list in self.ptema]
        self.P_Tema_kode.setCurrentIndex(self.ptema_list.index('324 - Landmålte punkt'))
        self.P_Tema_kode.setToolTip('Select PTema')
        

        self.mainLayout.addWidget(meth_lab, 4, 0)
        self.mainLayout.addWidget(self.method_code)
        self.method = [('10 - Uspesifisert målemetode', 10, 'Uspesifisert målemetode'),
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
                        ('966 - GPS-Rtk NC', 966, 'GPS-Rtk NC')]
        
        self.method_code.addItems([list[0] for list in self.method])
        self.method_list = [list[0] for list in self.method]
        self.method_code.setCurrentIndex(self.method_list.index('96 - GPS-Rtk'))
        self.method_code.setToolTip('Select målemetode')
        

        self.mainLayout.addWidget(noyakt_lab, 5, 0)
        self.mainLayout.addWidget(self.noyaktighet)
        self.noy = [('5', 5), 
                    ('10', 10), 
                    ('20', 20), 
                    ('30', 30), 
                    ('100', 100), 
                    ('200', 200), 
                    ('300', 300), 
                    ('400', 400), 
                    ('500', 500), 
                    ('1000', 1000)]
        
        self.noyaktighet.addItems([list[0] for list in self.noy])
        self.noy_list = [list[0] for list in self.noy]
        self.noyaktighet.setCurrentIndex(self.noy_list.index('5'))
        self.noyaktighet.setToolTip('Select noyaktighet')


        self.mainLayout.addWidget(syn_lab, 6, 0)
        self.mainLayout.addWidget(self.synbarhet)
        self.syn = [('0 - Fullt synlig ved innmåling', 0, 'Fullt synlig ved innmåling'), 
                    ('1 - Innmåling lukket grøft', 1, 'Innmåling lukket grøft'), 
                    ('3 - Ikke synlig trase; i sjø/undergrunn', 2, 'Ikke synlig trase; i sjø/undergrunn')]
        
        self.synbarhet.addItems([list[0] for list in self.syn])
        self.syn_list = [list[0] for list in self.syn]
        self.synbarhet.setCurrentIndex(self.syn_list.index('1 - Innmåling lukket grøft'))        
        self.synbarhet.setToolTip('Select Synbarhet')

        self.mainLayout.addWidget(okButton, 7, 1)
        self.mainLayout.addWidget(runButton, 7, 2)

    def selectInfile(self):
        inputFile = QFileDialog.getOpenFileName(self, "Finne målefila", '', "Text files (*.csv) ;; Excel files (*.xlsx) ;; KOF files (*.kof)")
        inputPath = Path(inputFile[0])
        self.input_stem = str(inputPath.stem)
        self.input_string = str(inputPath)
        self.inputFile.setText(self.input_string)
    
    def test_in(self):
        if not os.path.isfile(self.inputFile.text()):
            self.infile_error()

    def selectOutfile(self):
        outputFolder = QFileDialog.getExistingDirectory(self, 'Select Folder')
        outputPath = Path(outputFolder)
        self.outputFile.setText(str(outputPath))

    
    def get_outfile_name(self):      

        pattern = r'^[A-Za-z]{2}\d{6}$'
        if re.match(pattern, self.malernummer.text()):
           
            if self.outputFile.text().endswith(".shp") and os.path.isdir(os.path.dirname(self.outputFile.text())):
                self.outfile_path = os.path.dirname(self.outputFile.text())

            elif os.path.isdir(self.outputFile.text()):
                self.outfile_path = Path(self.outputFile.text())
                
                self.outfile_name = self.input_stem + '_' + self.malernummer.text()
                self.out_path = os.path.join(self.outfile_path, self.outfile_name)
                self.out_path = Path(self.out_path).with_suffix('.shp')
                self.outputFile.setText(str(self.out_path))
            
            elif not os.path.isdir(os.path.dirname(self.outputFile.text())):
                self.outfile_warn()
                self.auto_out()
        
        else:
            self.name_warn()
        
    
    def auto_out(self):
        self.outputFile.setText(os.path.dirname(self.inputFile.text()))
    
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
            'There was a problem with the input file, likely due to format of ".kof" file or differences of ".csv" file to VLoc3 version "2025 - VMMAP Web - 2.19.6" \n\n\n {}'.format(format_exc()),
            buttons=QMessageBox.StandardButton.Close
        )

    def complete(self):
         QMessageBox.information(
        self,
        "Shapefile Written",
        "Your shapefile has been written to {:}.".format(self.outputFile.text()),
        buttons=QMessageBox.StandardButton.Ok
    )

    
    def read_file(self):
        try:
            if Path(self.inputFile.text()).suffix in ['.kof', '.KOF']:
                ...
            elif Path(self.inputFile.text()).suffix in '.csv':
                
                with open(Path(self.inputFile.text())) as file:
                    dialect = csv.Sniffer().sniff(file.read(1024))
                    delim = dialect.delimiter
                    print(f'DELIMITER:   {delim}')
                    
                try:
                    csv_df = pd.read_csv(Path(self.inputFile.text()), sep = delim)
                except: 
                    csv_df = pd.read_excel(Path(self.inputFile.text()), sep = delim)

                csv_df.columns = (csv_df.columns.str.strip().str.lower().str.replace(" ", "").str.replace("[()€$]", "", regex=True))

                csv_df = csv_df.iloc[1:]
                csv_df = csv_df.copy()
                csv_df.reset_index(inplace=True)


                northing_label = csv_df.columns[csv_df.columns.str.contains('northing')][0]
                easting_label  = csv_df.columns[csv_df.columns.str.contains('easting')][0]
                altitude_label = csv_df.columns[csv_df.columns.str.contains('altitude/')][0]
                depth_label = csv_df.columns[csv_df.columns.str.contains('depth')][0]
                time_label = csv_df.columns[csv_df.columns.str.contains('gpstime')][0]
               
                try:
                    point_label = csv_df.columns[csv_df.columns.str.contains('up#')][0]
                except:
                    try:
                        point_label = csv_df.columns[csv_df.columns.str.contains('recordindex')][0]
                        csv_df[point_label] = 'UP' + csv_df[point_label].astype(int).astype(str)
                    except: 
                        csv_df['up#'] = 'UP' + csv_df.index.astype(str)
                        point_label = 'up#'

                fix_label = csv_df.columns[csv_df.columns.str.contains('gpsfix')][0]
                current_label = csv_df.columns[csv_df.columns.str.contains('locatecurrentma')][0]
                gain_label = csv_df.columns[csv_df.columns.str.contains('locatorgaindb')][0]
                bargraph_label = csv_df.columns[csv_df.columns.str.contains('bargraph%')][0]
                rms_label = csv_df.columns[csv_df.columns.str.contains('2drmsm')][0]




                csv_df = csv_df[csv_df[fix_label] != 'NONE']
                csv_df = csv_df.copy()
                csv_df.reset_index(inplace=True)

                if csv_df[altitude_label].dtype == 'O':
                    print(type(csv_df[altitude_label][0]))
                    csv_df[altitude_label] = pd.to_numeric(csv_df[altitude_label].str.replace(",", "."), errors="coerce")
                    csv_df[depth_label] = pd.to_numeric(csv_df[depth_label].str.replace(",", "."), errors="coerce")
                
                csv_df['cablealtitude'] = csv_df[altitude_label] - csv_df[depth_label]

                geometry = [Point(xyz) for xyz in zip(csv_df[easting_label], csv_df[northing_label], csv_df['cablealtitude'])]
                shp_gdf = gpd.GeoDataFrame(csv_df, geometry=geometry, crs="EPSG:25832")
                
                shp_gdf['KOORDH'] = csv_df['cablealtitude']
                shp_gdf['DATO'] = csv_df[time_label]
                shp_gdf['PNUMMER'] = csv_df[point_label]
                shp_gdf['GPS_fix'] = csv_df[fix_label]
                shp_gdf['Current_mA'] = csv_df[current_label]
                shp_gdf['Gain_dB'] = csv_df[ gain_label]
                shp_gdf['Bargraph_%'] = csv_df[bargraph_label]
                shp_gdf['2DRMS_m'] = csv_df[rms_label]

                shp_gdf['LANDMALER'] = self.malernummer.text().upper()[:5]
                shp_gdf['PTEMA'] = self.ptema[self.P_Tema_kode.currentIndex()][1]
                shp_gdf['TEMATEKST'] = self.ptema[self.P_Tema_kode.currentIndex()][3]
                shp_gdf['MALEMETODE'] = self.method[self.method_code.currentIndex()][1]
                shp_gdf['NOYAKTIGHE'] = self.noy[self.noyaktighet.currentIndex()][0]
                shp_gdf['SYNBARHET'] = self.syn[self.synbarhet.currentIndex()][1]  
                shp_gdf['H_MALEMETODE'] = self.method[self.method_code.currentIndex()][1]
                shp_gdf['H_NOYAKTIGHE'] = self.noy[self.noyaktighet.currentIndex()][0]

                print(shp_gdf.columns)
                
                shp_gdf = shp_gdf[['PNUMMER', 
                                   'DATO', 
                                   'LANDMALER', 
                                   'GPS_fix', 
                                   '2DRMS_m', 
                                   'Current_mA', 
                                   'Gain_dB', 
                                   'Bargraph_%', 
                                   'KOORDH', 
                                   'PTEMA', 
                                   'TEMATEKST',
                                   'MALEMETODE',
                                   'H_MALEMETODE', 
                                   'NOYAKTIGHE',
                                   'H_NOYAKTIGHE', 
                                   'SYNBARHET', 
                                   'geometry']].copy()
                
                schema = gpd.io.file.infer_schema(shp_gdf)

                for field in ['PTEMA']: #long
                    schema['properties'][field] = 'int32:10'

                for field in ['MALEMETODE', 'SYNBARHET', 'H_MALEMETODE']: #short
                    schema['properties'][field] = 'int32:4'

            
                
            shp_gdf.to_file(self.outputFile.text(), driver="ESRI Shapefile", schema=schema, engine="fiona")
            shp_gdf.to_csv(Path(self.outputFile.text()).with_suffix('.csv'), index=False)

            self.complete()

        except Exception:
            self.csv_error()
        
if __name__=="__main__":
    import sys
    app    = QApplication(sys.argv)
    myshow = InputDialog()
    myshow.setWindowTitle("CSV2SHP")
    app.setWindowIcon(QIcon('CSV2SHP.ico'))
    myshow.setWindowIcon(QIcon('CSV2SHP.ico'))
    myshow.show()
    app.exec()
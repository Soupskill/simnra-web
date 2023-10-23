import time
import os
import numpy as np

from win32com.client import Dispatch




class Simulation:

    def __init__(self, name: str) -> None:
        
        self.__app = Dispatch('Simnra.App')
        self.__setup = Dispatch('Simnra.Setup')
        self.__projectile = Dispatch('Simnra.Projectile')
        self.__taget = Dispatch('Simnra.Target')
        self.__cross_section = Dispatch('Simnra.CrossSec')
        self.__foil = Dispatch('Simnra.Foil')
        self.__spectrum = Dispatch('Simnra.Spectrum')

        self.name = name

    @property
    def name(self):

        return self.__app.OLEUser
    
    @name.setter
    def name(self, name: str):
        self.__app.OLEUser = name

    
    def calculate(self):
        return self.__app.CalculateSpectrum()
    
    def _show(self):
        self.__app.Show()

    def open(self, fname: str):
        """open xnra files"""
        return self.__app.Open(fname)
    
    def read_experimental(self, fname: str):

        if not os.path.isfile(fname):
            return False
        
        fileformats = {'ascii': 1,
                        'rbs': 14,
                        'mpa': 11,
                        'xnra': 9}
        extention = fname.split('.')[-1].lower()
        if extention not in fileformats.keys():
            extention = 'ascii'
        
        return self.__app.ReadSpectrumData(fname, fileformats[extention])
        



if __name__=="__main__":

    sim = Simulation('asd')
    sim.read_experimental('C:/Users/CHE/Desktop/Prusachenko/09-20-2023/R2_PLX_E=500keV_Q=25.rbs')
    print(sim.name)
    sim._show()
    time.sleep(5)

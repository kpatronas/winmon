#!/usr/bin/env python

import argparse
import getpass
import ctypes
import json
import time
import wmi
import sys
import os
 
from time        import mktime,strptime
from datetime    import date,datetime,timedelta
from collections import defaultdict
from tabulate    import tabulate


def bytesto(bytes, to, bsize=1024):
    a = {'k' : 1, 'm': 2, 'g' : 3, 't' : 4, 'p' : 5, 'e' : 6 }
    r = float(bytes)
    for i in range(a[to]):
        r = int(r / bsize)
    return(r)
    
def windowsDateConversion(date, displayHour = True):
    if (len(date)==8 or len(date)==25):
        finaldate = date[6:8]+"/"+date[4:6]+"/"+date[0:4]
        if displayHour:
            finaldate = finaldate +" "+date[8:10]+":"+date[10:12]+":"+date[12:14]
    else:   
        finaldate = date
    return finaldate
    
class WindowsComputer:
  
    def __init__(self, hostname = None, username = None, password = None):
        # Arguments
        self.hostname = hostname
        self.username = username
        self.password = password 
          
        self.computer                       = None
        self.generic_os_info                = defaultdict(dict)
        self.ram_usage                      = defaultdict(dict)
        self.cpu_usage                      = defaultdict(dict)
        self.disk_usage                     = defaultdict(dict)
        self.automatic_services_not_running = defaultdict(dict)
        self.network_info                   = defaultdict(dict)
        self.network_speed                  = defaultdict(dict)
        
        self.generic_os_info_tbl                = []
        self.ram_usage_tbl                      = []
        self.cpu_usage_tbl                      = []
        self.disk_usage_tbl                     = []
        self.automatic_services_not_running_tbl = []
        self.network_info_tbl                   = []
        self.network_speed_tbl                  = []
          
          
    def wmi_connect(self):
        # If user name is given but password is not, ask for a password
        if self.username and not self.password:
            self.password = getpass.getpass(prompt='Password: ')
        # Try to connect        
        try:
            if self.username:
                self.computer = wmi.WMI(computer = self.hostname, user = self.username, password = self.password)
                return True
            else:
                self.computer = wmi.WMI(computer = self.hostname)
                return True
        except wmi.x_wmi:
            print "Could not connect to %s " %self.hostname
            return False
        
    def get_os_info(self):
        obj_property = None
        for obj_property in self.computer.query("SELECT * FROM Win32_OperatingSystem"):
            self.generic_os_info['os']['Name']          = obj_property.Caption                                # Get the OS name
            self.generic_os_info['os']['CSName']        = obj_property.CSName                                 # Get the Server name
            self.generic_os_info['os']['arch']          = obj_property.OSArchitecture                         # Get the OS Architecture
            self.generic_os_info['os']['BootUpTime']    = windowsDateConversion(obj_property.LastBootUpTime)  # Get the OS Last Boot Up Time
            self.generic_os_info['os']['InstallDate']   = windowsDateConversion(obj_property.InstallDate)     # Get the OS Install Date
            self.generic_os_info['os']['LocalDateTime'] = windowsDateConversion(obj_property.LocalDateTime)   # Get the OS Local Date Time
            self.generic_os_info['os']['SrvPackMjrVer'] = str(obj_property.ServicePackMajorVersion)           # Get the OS Service Pack Major version
            self.generic_os_info['os']['SrvPackMinVer'] = str(obj_property.ServicePackMinorVersion)           # Get the OS Service Pack Minor version
   
            self.generic_os_info_tbl.append(["OS","Host","Arch","Boot Up Time","Install Date","Local Date Time","Srv Pack Version"])
            self.generic_os_info_tbl.append([self.generic_os_info['os']['Name'],
                                             self.generic_os_info['os']['CSName'],
                                             self.generic_os_info['os']['arch'],
                                             self.generic_os_info['os']['BootUpTime'],
                                             self.generic_os_info['os']['InstallDate'],
                                             self.generic_os_info['os']['LocalDateTime'],
                                             self.generic_os_info['os']['SrvPackMjrVer']+"."+self.generic_os_info['os']['SrvPackMinVer']])

        return tabulate(self.generic_os_info_tbl,headers="firstrow",numalign="center")
              
    def get_memory_usage(self):
        obj_property = None
        for obj_property in self.computer.query("SELECT * FROM Win32_OperatingSystem"):
            self.ram_usage['ram']['FreePhysicalMemory']           = int(obj_property.FreePhysicalMemory)/1000                                          
            self.ram_usage['ram']['TotalVisibleMemorySize']       = int(obj_property.TotalVisibleMemorySize)/1000                                          
            self.ram_usage['ram']['UsedPhysicalMemory']           = int(obj_property.TotalVisibleMemorySize)/1000-(int(obj_property.FreePhysicalMemory)/1000) 
            self.ram_usage['ram']['FreePhysicalMemoryPercentage'] = (int(obj_property.FreePhysicalMemory)*100) / int(obj_property.TotalVisibleMemorySize)     
            self.ram_usage['ram']['UsedPhysicalMemoryPercentage'] = 100-int(self.ram_usage['ram']['FreePhysicalMemoryPercentage'])                            
        
            self.ram_usage_tbl.append(["Total RAM (MB)","Free Ram (MB)","Used Ram (MB)","Free Ram (%)","Used Ram (%)"])
            self.ram_usage_tbl.append([self.ram_usage['ram']['TotalVisibleMemorySize'],
                                       self.ram_usage['ram']['FreePhysicalMemory'],
                                       self.ram_usage['ram']['UsedPhysicalMemory'],
                                       self.ram_usage['ram']['FreePhysicalMemoryPercentage'],
                                       self.ram_usage['ram']['UsedPhysicalMemoryPercentage']])
        return tabulate(self.ram_usage_tbl,headers="firstrow",numalign="center")
              
    def get_cpu_usage(self):
        obj_property = None
        for obj_property in self.computer.query("SELECT * FROM win32_processor"):
            self.cpu_usage['cpu']['LoadPercentage'] = obj_property.LoadPercentage 
            self.cpu_usage['cpu']['Name']           = obj_property.Name           
            self.cpu_usage['cpu']['NumberOfCores']  = obj_property.NumberOfCores
            self.cpu_usage_tbl.append(["CPU Load %","Type","Cores"])
            self.cpu_usage_tbl.append([self.cpu_usage['cpu']['LoadPercentage'],self.cpu_usage['cpu']['Name'],self.cpu_usage['cpu']['NumberOfCores']])

        return tabulate(self.cpu_usage_tbl,headers="firstrow",numalign="center")

              
    def get_disk_usage(self):
        obj_property = None
        disks        = []
        self.disk_usage_tbl.append(["Drive","Size (GB)","Free Space (GB)","Free Space (%)","Used Space (%)"])
        for obj_property in self.computer.query("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3"):
            self.disk_usage['disk']['Caption']             = obj_property.Caption                                       # Get Disk Caption
            self.disk_usage['disk']['FreeSpace']           = bytesto(int(obj_property.FreeSpace),'g') or 0              # Get Disk Free Space
            self.disk_usage['disk']['Size']                = bytesto(int(obj_property.Size),'g')      or 0              # Get Disk Used Space
            self.disk_usage['disk']['FreeSpacePercentage'] = (int(obj_property.FreeSpace)*100) / int(obj_property.Size) # Get the percentage of Free memory
            self.disk_usage['disk']['UsedSpacePercentage'] = (100-self.disk_usage['disk']['FreeSpacePercentage'])       # Get the percentage of Used memory
            self.disk_usage_tbl.append([self.disk_usage['disk']['Caption'],self.disk_usage['disk']['Size'],self.disk_usage['disk']['FreeSpace'],self.disk_usage['disk']['FreeSpacePercentage'],self.disk_usage['disk']['UsedSpacePercentage']])
            disks.append(dict(self.disk_usage))
  
        return tabulate(self.disk_usage_tbl,headers="firstrow",numalign="center")
        
    def get_automatic_services_not_running(self):
        obj_property = None
        self.automatic_services_not_running_tbl.append(["Automatic Services Not Running","Name"])
        obj_property = self.computer.query("SELECT * FROM Win32_Service where State='Stopped' and StartMode='Auto'")
        if obj_property:
            for service in obj_property:
                self.automatic_services_not_running['service']['Caption'] = service.caption
                self.automatic_services_not_running['service']['Name']    = service.Name
                self.automatic_services_not_running_tbl.append([self.automatic_services_not_running['service']['Caption'],self.automatic_services_not_running['service']['Name']])
            return tabulate(self.automatic_services_not_running_tbl,headers="firstrow",numalign="center")
        else:
            print "No auto services stopped"
            
    def get_network_info(self):
        obj_property  = None
        self.network_info_tbl.append(["Interface","DHCP Enabled","IPv4","IPv6","MAC"])
        for obj_property in self.computer.query("SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled = True"):
            self.network_info['Network']['Description'] = obj_property.Description
            self.network_info['Network']['DHCPEnabled'] = "True" if obj_property.DHCPEnabled == 1 else "False"
            self.network_info['Network']['IPv4']        = obj_property.IPAddress[0]
            self.network_info['Network']['IPv6']        = obj_property.IPAddress[1]
            self.network_info['Network']['MACAddress']  = obj_property.MACAddress
            self.network_info_tbl.append([self.network_info['Network']['Description'],self.network_info['Network']['DHCPEnabled'],self.network_info['Network']['IPv4'],self.network_info['Network']['IPv6'],self.network_info['Network']['MACAddress']])
        return tabulate(self.network_info_tbl,headers="firstrow",numalign="center")
                        
if __name__ == '__main__':
      
    obj_parser = argparse.ArgumentParser(description='Get a windows host system metrics using WMI')
    obj_parser.add_argument('-s', '--server', help = 'Host to query',          required = False, default = os.environ['COMPUTERNAME'])
    obj_parser.add_argument('-u', '--user',   help = 'username',               required = False)
    obj_parser.add_argument('-p', '--pass',   help = 'password',               required = False)
      
    d_args       = vars(obj_parser.parse_args())  # Validate and get arguments
    obj_computer = WindowsComputer(hostname = d_args['server'], username = d_args['user'], password =  d_args['pass'])
    
    if obj_computer.wmi_connect():
        print obj_computer.get_os_info()+"\n"
        print obj_computer.get_cpu_usage()+"\n"
        print obj_computer.get_memory_usage()+"\n"
        print obj_computer.get_disk_usage()+"\n"
        print obj_computer.get_automatic_services_not_running()+"\n"
        print obj_computer.get_network_info()+"\n"
    else:
        sys.exit()

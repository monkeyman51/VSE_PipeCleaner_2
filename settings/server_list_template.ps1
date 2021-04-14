#=================================================================================================================================
# Copyright (c) Microsoft Corporation
# All rights reserved. 
# MIT License
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files 
# (the ""Software""), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, 
# merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is 
# furnished to do so, subject to the following conditions:
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES 
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
# LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF 
# OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#=================================================================================================================================

#---------------------------------------------------------------------------------------------------
# Short description (less than 80 char) that describes this group of servers and WCS manager
#---------------------------------------------------------------------------------------------------
$ServerListDescription = "Broadwell servers for NVMe qualification"

$ServerList_NvmeOobSupport                = $FALSE
$ServerList_NvmeTempSensorSupport         = $FALSE
 
# The following settings can be used to change the power cycle on/off times for the servers in this server list file
# If NVDIMM-SW is enabled, set SETTING_POWER_CYCLE_OFFTIME to 600
#------------------------------------------------------------------------------------------------------------------------------------
$SETTING_POWER_CYCLE_ONTIME               = 1800           # For WCS based power cycles, timeout between power on and check for results
$SETTING_POWER_CYCLE_OFFTIME              = 30             # For WCS based power cycles, time between power off and on in seconds 

#---------------------------------------------------------------------------------------------------
# Enter server names and locations as hash tables in the PowerShell array: $WcsTestRemoteBladeInfo
#---------------------------------------------------------------------------------------------------
# Hash table should be in this format: @{ slot = 13;   Address = '172.16.0.210'  ; Name = 'WCSMTS13QPVT001'}
#
# Slot # is the WCS server location, if not using a WCS managed chassis or rack enter an
# incrementing number for each slot.  Example: Slot = 1, Slot = 2, etc.
#
# Hostname can be either the server hostname or IPV4 address
#
# For platforms that come with JBOD/JBOF (J2010/F2010), each server should also contain a corresponding JBOD/JBOF slot #.
# Example for J2010/F2010, where up to 2 blades can be connected to the same storage enclosure: 
#     @{ slot = 30;   Address = '10.177.226.69'  ; Name = 'CSIG6WAQLPRD057' ; PerifSlot = 32 }
#     @{ slot = 31;   Address = '10.177.226.135' ; Name = 'CSIG6WAQLPRD058' ; PerifSlot = 32 }
# Example for platform that has 1 blade connected to 2 different JBODs:
#     @{ slot = 32;   Address = '10.218.112.122' ; Name = 'CSI7ZT9228DV005' ; PerifSlot = @(27,33) }
#
# For G50 platforms each server should also contain a corresponding G50 slot #.
# Example:
#     @{ slot = 30;   Address = '10.177.226.69'  ; Name = 'CSIG6WAQLPRD057' ; G50Slot = 29 }
#
# For Celestial Peak SKUs, each server should contain a corresponding SOC IP.
# Example:
#    @{ slot = 1;   Address = '10.177.237.37'  ; Name = 'CSI7W3506EVT016'; SOCIp = '10.177.236.90'}
#
#---------------------------------------------------------------------------------------------------
$WcsTestRemoteBladeInfo= @(

# {{insert_server_list}}

)
#---------------------------------------------------------------------------------------------------
# Enter the rack or chassis manager type.   Allowed versions are : 
#---------------------------------------------------------------------------------------------------
#
#   'None'    to indicate no WCS manager used
#   'Legacy'  to indicate WCS chassis manager used (gen4/5 compute and storage)
#   'M2010'   to indicate WCS rack manager for gen6 
#   'RMM'     to indicate stand-alone WCS rack manager for gen6
#   'G50'     to indicate stand-alone WCS rack manager for gen6 with PCIe expander
#---------------------------------------------------------------------------------------------------
$WcsTestRemoteMgrType            = # {{insert_type}}

#---------------------------------------------------------------------------------------------------
# Enter hostname or IPV4 address of the rack or chassis manager. If not used enter blank string:  ''
#---------------------------------------------------------------------------------------------------
$WcsTestRemoteMgr                = # {{insert_server_dhcp}}

#---------------------------------------------------------------------------------------------------
# If using legacy chassis manager enter $true if using SSL for REST communication.   
# If not using a legacy chassis manager or SSL is disabled enter $false
#---------------------------------------------------------------------------------------------------
$WcsTestRemoteMgrSslEnabled      =  $true


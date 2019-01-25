# -*- coding: utf-8 -*-

__author__ = "Jason Smith"
__credits__ = []
__email__ = "jasons2@cisco.com"

"""
Base template for Python Scripts.
Created by Jason Smith (jasons2@cisco.com)  
"""

### TEMPLATE IMPORTS
import logging
import logging.config
import yaml
import argparse
import datetime

### IMPORTS
import requests
import json
import xlsxwriter

## TEMPLATE FUNCTIONS
def setupLogging(logconfig, logfile, verbose, debug):
    """
    Setup logging configuration
    """
    global config

    with open(logconfig, 'rt') as f:
        config = yaml.safe_load(f.read())

    if logfile:
        configureLogging(handler='file_handler', filename=logfile)

    if verbose:
        configureLogging(handler='file_handler', formatter='verbose')
        configureLogging(handler='console', formatter='verbose')

    if debug:
        configureLogging(handler='file_handler', level='DEBUG')
        configureLogging(handler='console', level='DEBUG')

    logging.config.dictConfig(config)

def configureLogging(handler='', formatter='', level='', filename=''):
    """
    """
    global config, logger

    if formatter:
        config['handlers'][handler]['formatter'] = formatter
    if level:
        config['handlers'][handler]['level'] = level

    if filename and handler =='file_handler':
        config['handlers'][handler]['filename'] = filename

    logging.config.dictConfig(config)

def getArgs():
    """
    Get Arguments from Command Line.  Set defaults for logfile, logconfig, debug and verbose.
    """
    parser = argparse.ArgumentParser()
    parser.set_defaults(logfile='',
                        logconfig = 'logging.yaml',
                        debug=False,
                        verbose=False)

    # Default flags for arguments with defaults set.
    parser.add_argument('-l', '--logfile', help='Logging Filename (default is script_name.log)')
    parser.add_argument('-c', '--logconfig', help='Logging Configuration File (default is logging.yaml)')
    parser.add_argument('-d', '--debug', help='Set Debug', dest='debug', action='store_true')
    parser.add_argument('-v', '--verbose', help='Set Logging to Verbose', dest='verbose', action='store_true')
    parser.add_argument('-n', '--networkinfo', help='Network Name or Network Id')
    parser.add_argument('-o', '--orginfo', help='Org Id or Org Name')
    parser.add_argument('-a', '--apikey', help='Set API Key', required=True)
    parser.add_argument('-f', '--filename', help='Filename')

    # Add additional arguments if required
    #parser.add_argument('-e', '--example', help='This is an example')

    output = parser.parse_args()  # assign contents of parser to output.

    if not output.logfile:  # Check to see if there is a logfile in command line.
        # Assign logfile to name of script plus .log
        output.logfile = "_" + __file__.split(".")[0] + ".log"

    return output

def info(msg):
    logger.info(msg)

def error(msg):
    logger.error(msg)

def warning(msg):
    logger.warning(msg)

def debug(msg):
    logger.debug(msg)

### HELPER FUNCTIONS
def getOrganizations(f_key):
    info("Gathering Oganizations")
    url = 'https://dashboard.meraki.com/api/v0/organizations'
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        if r.status_code != requests.codes.ok:
            return 'null'

        return r.json()

    except Exception as e:
        error(e)

def getOrganizationId(f_key, f_orginput):
    info("Determining Oganization ID")
    url = 'https://dashboard.meraki.com/api/v0/organizations'
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        if r.status_code != requests.codes.ok:
            return 'null'

        rjson = r.json()

        try:
            for record in rjson:
                if record['id'] == int(f_orginput):
                  return record['id']
        except:
            for record in rjson:
                if record['name'] == f_orginput:
                  return record['id']

        return('null')

    except Exception as e:
        error(e)

def getOrganizationName(f_key, f_orginput):
    info("Determining Oganization Name")
    url = 'https://dashboard.meraki.com/api/v0/organizations'
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        if r.status_code != requests.codes.ok:
            return 'null'

        rjson = r.json()

        try:
            for record in rjson:
                if record['id'] == int(f_orginput):
                  return record['name']
        except:
            for record in rjson:
                if record['name'] == f_orginput:
                  return record['name']

        return('null')

    except Exception as e:
        error(e)

def getShardURL(f_key, f_orgid):
    info("Gathering Shard URL")
    try:
        url = 'https://dashboard.meraki.com/api/v0/organizations/%s/snmp' % f_orgid
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        if r.status_code != requests.codes.ok:
            return 'null'

        rjson = r.json()

        return(rjson['hostname'])
    except Exception as e:
        error(e)

def getNetworks(f_key, f_shardurl, f_orgid):
    info("Gathering Networks")
    try:
        url = 'https://%s/api/v0/organizations/%s/networks' % (f_shardurl, f_orgid)
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        returnvalue = []
        if r.status_code != requests.codes.ok:
            returnvalue.append({'name': 'null', 'id': 'null'})
            return(returnvalue)

        return(r.json())
    except Exception as e:
        error(e)

def getNetworkId(f_key, f_shardurl, f_nwInfo, f_orgid):
    info("Determining Network Id")
    try:
        url = 'https://%s/api/v0/organizations/%s/networks' % (f_shardurl, f_orgid)

        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        if r.status_code != requests.codes.ok:
            return 'null'

        rjson = r.json()

        for r in rjson:
            if r['name'] == f_nwInfo:
                return r['id']
            elif r['id'] == f_nwInfo:
                return r['id']
        return 'null'

    except Exception as e:
        error(e)

def getSwitchPorts(f_key, f_serialnumber, f_shardurl):
    info("Gathering Individual Switchport Configurations")
    try:
        url = 'https://%s/api/v0/devices/%s/switchPorts' % (f_shardurl, f_serialnumber)
      
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})
        
        if r.status_code != requests.codes.ok:
            return {'Return Code': r.status_code}
        return (r.json())

    except Exception as e:
        error((e, r.status_code))

def getSSIDs(f_key, f_shardurl, f_networkid):
    info("Gathering SSID information")
    url = 'https://%s/api/v0/networks/%s/ssids' % (f_shardurl, f_networkid)

    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})

        if r.status_code != requests.codes.ok:
            return {'Return Code':r.status_code}

        return r.json()
    except Exception as e:
        error(e)

def getDevices(f_key, f_shardurl, f_networkid):
    info("Gathering Device Information")
    url = 'https://%s/api/v0/networks/%s/devices' % (f_shardurl, f_networkid)
    
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_key, 'Content-Type': 'application/json'})
        
        if r.status_code != requests.codes.ok:
            print r.status_code
            return {'Return Code': r.status_code}
        
        return r.json()
    
    except Exception as e:
        error((e, r.status_code))

def getStaticRoutes(f_apikey, f_shardurl, f_networkid):
    info("Gathering Static Route Information")
    url = 'https://%s/api/v0/networks/%s/staticRoutes' % (f_shardurl, f_networkid)
    
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_apikey, 'Content-Type': 'application/json'})
        
        if r.status_code != requests.codes.ok:
            return {'Return Code': r.status_code}
        print (r.json())
        return r.json()
    
    except Exception as e:
        error(e)

def getVlans(f_apikey, f_shardurl, f_networkid):
    info("Gathering VLAN Information")
    url = 'https://%s/api/v0/networks/%s/vlans' % (f_shardurl, f_networkid)
    
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_apikey, 'Content-Type': 'application/json'})
        
        if r.status_code != requests.codes.ok:
            if r.status_code == 400:
                return []
            else:
                return {'Return Code': r.status_code}
        
        return r.json()
    
    except Exception as e:
        error(e)

def getSwitchSettings(f_apikey, f_shardurl, f_networkid):
    info("Gathering Switch Configuration")
    url = 'https://%s/api/v0/networks/%s/switch/settings' % (f_shardurl, f_networkid)
    
    try:
        r = requests.get(url, headers={'X-Cisco-Meraki-API-Key': f_apikey, 'Content-Type': 'application/json'})
        
        if r.status_code != requests.codes.ok:
            return {'Return Code': r.status_code}
        
        return r.json()
    
    except Exception as e:
        error(e)

def createXL(f_fname, f_ssidList, f_swPorts, f_deviceList, f_vlanList):

    wb = xlsxwriter.Workbook(f_fname)

    # Create Devices Tab
    info("Creating and Populating Device Tab")
    deviceWS = wb.add_worksheet('Devices')
    deviceColTitles = ['networkId',
                       'model',
                       'name',
                       'tags',
                       'lanIp',
                       'address',
                       'lat',
                       'serial',
                       'lng']
    deviceWS.write_row(0,0,deviceColTitles)
    row = 1
    for _dev in f_deviceList:
        try:
            deviceWS.write_row(row,0,_dev)
        except Exception as e:
            error(e)
        row += 1
    info("Device Tab Complete")

    # Create SSIDS Tab
    info("Creating and Populating SSID Tab")
    ssidsWS = wb.add_worksheet('SSIDS')
    ssidsColTitles = ['name',
                      'splashPage',
                      'perClientBandwidthLimitDown',
                      'enabled',
                      'number',
                      'perClientBandwidthLimitUp',
                      'ssidAdminAccessible',
                      'bandSelection',
                      'minBitrate',
                      'ipAssignmentMode',
                      'authMode',
                      'wpaEncryptionMode',
                      'useVlanTagging',
                      'radiusFailoverPolicy',
                      'radiusCoaEnabled',
                      'radiusAttributeForGroupPolicies',
                      'radiusServers',
                      'radiusOverride',
                      'radiusAccountingEnabled',
                      'radiusLoadBalancingPolicy',
                      'encryptionMode']

    ssidsWS.write_row(0,0,ssidsColTitles)
    row = 1
    for _ssid in f_ssidList:
        try:
            ssidsWS.write_row(row,0,_ssid)
        except Exception as e:
            error((e, _ssid))
        row += 1
    info("SSID Tab Complete")

    # Create Switch Ports Tab
    for switchName, swDetails in f_swPorts.iteritems():
        info("Creating and Populating %s Tab" % (switchName))
        _tempSheet = wb.add_worksheet(switchName)
        _tempSheet.write_row(0,0,swDetails['headings'])
        row = 1
        for _swPortDetails in swDetails['ports']:
            _tempSheet.write_row(row,0,_swPortDetails)
            row += 1
        info("%s Tab Complete" % (switchName))

    # Create VLAN Tab
    info("Creating and Populating VLANS Tab")
    vlanWS = wb.add_worksheet('VLANS')
    vlanColTitles = ['networkId', 
                     'subnet', 
                     'fixedIpAssignments', 
                     'name', 
                     'applianceIp', 
                     'reservedIpRanges', 
                     'dnsNameservers', 
                     'id']

    vlanWS.write_row(0,0,vlanColTitles)
    row = 1
    for _vlan in f_vlanList:
        try:
            vlanWS.write_row(row,0,_vlan)
        except Exception as e:
            error(e)
        row += 1
    info("VLANS Tab Complete")

    wb.close()

### MAIN FUNCTION
def main():
    """
    Main Function
    """
    info("== (" + scriptStartTime + ") START SCRIPT ==")
    # Set Arguments    
    args = getArgs()

    # Assign API key from commmand line arguments
    apikey = args.apikey

    # Get Organization ID based on information provided by user
    # User gives either OrgID or OrgName and the program finds
    # OrgID to use for other API calls
    if args.orginfo:
        orgid = getOrganizationId(apikey, args.orginfo) # Get Organization ID
        orgName = getOrganizationName(apikey, args.orginfo) # Get Organization Name
    else:
        info("Must provide either Org Id or Organization Name.")
        info("Please use the list below and specify on command line with -o")

        # Get List of Organizations that the user has access to via the API
        for _org in getOrganizations(apikey):
            info("Name \"%s\"  Id \"%s\" " % (_org['name'], _org['id']))

        scriptEndTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        info("== (" + scriptEndTime + ") END SCRIPT ==")
        exit()
    
    # Assign Filename from command line arguments
    # Append xlsx to the end of a filename that doesn't
    # have the excel extension.
    if args.filename:
        if '.xlsx' in args.filename:
            fname = args.filename
        else:
            fname = args.filename + '.xlsx'
    else:
        fname = orgName.replace(" ", "_") + "_" + datetime.datetime.now().strftime("%m%d%y_%H%M%S") + ".xlsx"

    # Get ShardURL.  This is a requirement in Meraki to ensure API calls are 
    # made to the right shard (database server instance).
    shardurl = getShardURL(apikey, orgid)
    debug ("getShardURL returned %s" % shardurl)

    # Get Network ID based on information provided by user
    # User gives either Network ID or Network Name and program finds
    # NetID to use for other API calls
    if args.networkinfo:
        netId = getNetworkId(apikey, shardurl, args.networkinfo, orgid)
    else:
        info("Must provide either Network Id or Network Name.")
        info("Please use the list below and specify on command line with -n")

        # Get List of Networks that the user has access to via the API
        for _net in getNetworks(apikey, shardurl, orgid):
            info("Name \"%s\"  Id \"%s\" " % (_net['name'], _net['id']))

        scriptEndTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        info("== (" + scriptEndTime + ") END SCRIPT ==")
        exit()

    # Debug statements if something isn't working right.  Specify -d in command line
    debug("Arguments are : args.apikey %s " % args.apikey)
    debug("Returned was : apikey %s " % apikey )
    debug("Arguments are : args.orginfo %s " % args.orginfo)
    debug("Returned was : orgid %s " % orgid )
    debug("Arguments are : args.networkinfo %s " % args.networkinfo)
    debug("Returned was : netId %s " % netId )
    
    # Gather Device Information
    devList = []
    for _dev in getDevices(apikey, shardurl, netId):
	print _dev
        devList.append([_dev['networkId'],
                        _dev['model'],
                        _dev['name'],
                        _dev['tags'] if 'tags' in _dev.keys() else 'NA',
                        _dev['lanIp'],
                        _dev['mac'],
                        _dev['address'],
                        _dev['lat'],
                        _dev['serial'],
                        _dev['lng']]
                       )

    # Gather SSID Information
    ssidList = []
    for ss in getSSIDs(apikey, shardurl, netId):
        ssidList.append([ss['name'],
                         ss['splashPage'],
                         ss['perClientBandwidthLimitDown'],
                         ss['enabled'],
                         ss['number'],
                         ss['perClientBandwidthLimitUp'],
                         ss['ssidAdminAccessible'],
                         ss['bandSelection'],
                         ss['minBitrate'],
                         ss['ipAssignmentMode'],
                         ss['authMode'],
                         ss['wpaEncryptionMode'] if 'wpaEncryptionMode' in ss.keys() else 'NA',
                         ss['useVlanTagging'] if 'useVlanTagging' in ss.keys() else 'NA',
                         ss['radiusFailoverPolicy'] if 'radiusFailoverPolicy' in ss.keys() else 'NA',
                         ss['radiusCoaEnabled'] if 'radiusCoaEnabled' in ss.keys() else 'NA',
                         ss['radiusAttributeForGroupPolicies'] if 'radiusAttributeForGroupPolicies' in ss.keys() else 'NA',
                         str(ss['radiusServers']) if 'radiusServers' in ss.keys() else 'NA',
                         ss['radiusOverride'] if 'radiusOverride' in ss.keys() else 'NA',
                         ss['radiusAccountingEnabled'] if 'radiusAccountingEnabled' in ss.keys() else 'NA',
                         ss['radiusLoadBalancingPolicy'] if 'radiusLoadBalancingPolicy' in ss.keys() else 'NA',
                         ss['encryptionMode'] if 'encryptionMode' in ss.keys() else 'NA']
                       )
    
    # Gather Switch Port information
    switchPortsDict = {}
    for _dev in devList:                     # Iterate over Device List
        if _dev[1][:2] == 'MS':              # Check to see that the device is a switch   
            _key = _dev[2]
            switchPortsDict[_key] = {'headings' : ['number',
                                                                             'name',
                                                                             'enabled',
                                                                             'allowedVlans',
                                                                             'voiceVlan',
                                                                             'vlan',
                                                                             'rstpEnabled',
                                                                             'linkNegotiation',
                                                                             'accessPolicyNumber',
                                                                             'stpGuard',
                                                                             'poeEnabled',
                                                                             'isolationEnabled',
                                                                             'type',
                                                                             'tags'],
                                                          'ports' : []}

            for sp in getSwitchPorts(apikey, _dev[8], shardurl):
                switchPortsDict[_key]['ports'].append([sp['number'],
                                                          sp['name'],
                                                          sp['enabled'],
                                                          sp['allowedVlans'],
                                                          sp['voiceVlan'],
                                                          sp['vlan'],
                                                          sp['rstpEnabled'],
                                                          sp['linkNegotiation'],
                                                          sp['accessPolicyNumber'],
                                                          sp['stpGuard'],
                                                          sp['poeEnabled'],
                                                          sp['isolationEnabled'],
                                                          sp['type'],
                                                          sp['tags']]
                                                          )

    # Gather VLAN information
    vlanList = []
    for _vlan in getVlans(apikey, shardurl, netId):
        vlanList.append([_vlan['networkId'], 
                         _vlan['subnet'], 
                         str(_vlan['fixedIpAssignments']), 
                         _vlan['name'], 
                         _vlan['applianceIp'], 
                         str(_vlan['reservedIpRanges']), 
                         _vlan['dnsNameservers'], 
                         _vlan['id']
                        ])

    # Create Excel document with gathered information
    info("Creating Excel Speadseet - \"%s\"" % (fname))
    createXL(fname, ssidList, switchPortsDict, devList, vlanList)
    info("Excel Speadseet - \"%s\" Complete" % (fname))

    scriptEndTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    info("== (" + scriptEndTime + ") END SCRIPT ==")

### MAIN PROGRAM
if __name__ == "__main__":
    # Record start time of script (after parser and compiler)
    scriptStartTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Create initial logger. It will only log to console and will be disabled
    # when we read the logging configuration from the config file.
    logging.basicConfig(level="INFO",
                        format='%(asctime)s [%(levelname)s] : %(name)s : %(message)s',
                        datefmt='%D %I:%H:%M')
    logger = logging.getLogger()

    # If you want to configure logging for External Modules
    # Example of Configuring Logging Levels for External Modules
    # mlogger = logging.getLogger('mimir')
    # mlogger.setLevel(logging.INFO)

    # Get arguments from command line.
    args = getArgs()

    # Setup default logging based on default config file.  Previous Logging config is overwritten
    setupLogging(logconfig=args.logconfig, logfile=args.logfile, verbose=args.verbose, debug=args.debug)

    # Output arguments to log
    debug("==ARGUMENTS==")
    debug("logfile : %s" % (args.logfile))
    debug("logconfig: %s" % (args.logconfig))
    debug("debug: %s" % (args.debug))
    debug("verbose: %s" % (args.verbose))
    debug("==END ARGUMENTS==")
    # Add additional info if you add additional arguments above
    # debug("example: %s" % (args.example))

    main()

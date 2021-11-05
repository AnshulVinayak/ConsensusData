# SimpleHistoryExample.py
from __future__ import print_function
from __future__ import absolute_import

from optparse import OptionParser

import os
import platform as plat
import sys
if sys.version_info >= (3, 8) and plat.system().lower() == "windows":
    # pylint: disable=no-member
    with os.add_dll_directory(os.getenv('BLPAPI_LIBDIR')):
        import blpapi
else:
    import blpapi

def parseCmdLine():
    parser = OptionParser(description="Retrieve reference data.")
    parser.add_option("-a",
                      "--ip",
                      dest="host",
                      help="server name or IP (default: %default)",
                      metavar="ipAddress",
                      default="localhost")
    parser.add_option("-p",
                      dest="port",
                      type="int",
                      help="server port (default: %default)",
                      metavar="tcpPort",
                      default=8194)

    options,_ = parser.parse_args()

    return options


def getData(ticker, data_date):
    options = parseCmdLine()

    # Fill SessionOptions
    sessionOptions = blpapi.SessionOptions()
    sessionOptions.setServerHost(options.host)
    sessionOptions.setServerPort(options.port)

    # print("Connecting to %s:%s" % (options.host, options.port))
    # Create a Session
    session = blpapi.Session(sessionOptions)

    # Start a Session
    if not session.start():
        print("Failed to start session.")
        return

    try:
        # Open service to get historical data from
        if not session.openService("//blp/refdata"):
            print("Failed to open //blp/refdata")
            return

        # Obtain previously opened service
        refDataService = session.getService("//blp/refdata")

        # Create and fill the request for the historical data
        request = refDataService.createRequest("HistoricalDataRequest")
        for t in ticker:
        # request.getElement("securities").appendValue("MSFT US Equity")
        #     print(t, data_date)
            request.getElement("securities").appendValue(t)
        request.getElement("fields").appendValue("cur_mkt_val")
        request.getElement("fields").appendValue("px_last")
        # request.getElement("fields").appendValue("OPEN")
        request.set("periodicityAdjustment", "ACTUAL")
        request.set("periodicitySelection", "DAILY")
        request.set("startDate", data_date)
        request.set("endDate", data_date)
        request.set("maxDataPoints", 1)

        # print("Sending Request:", request)
        # Send the request
        session.sendRequest(request)

        # Process received events
        msgs = []
        while(True):
            # We provide timeout to give the chance for Ctrl+C handling:
            ev = session.nextEvent(500)
            for msg in ev:
                msgs.append(msg)


            if ev.eventType() == blpapi.Event.RESPONSE:
                # Response completly received, so we could exit
                break
    finally:
        # Stop the session
        session.stop()
    vals = []
    # print(len(msgs))
    for idx in range(len(msgs)-3):
        try:
            # print(msgs[3])
            securities = msgs[idx+3].asElement().getElement('securityData')
            val1 = securities.getElement('fieldData').getValue().getElement('cur_mkt_val').getValue()
            val2 = securities.getElement('fieldData').getValue().getElement('px_last').getValue()
            sec = securities.getElement('security').getValue()
            # print({'name': sec, 'cur_mkt_val': val1})
            vals.append({'name': sec, 'cur_mkt_val':val1, 'px_last':val2})
        except:
            securities = msgs[idx + 3].asElement().getElement('securityData')
            sec = securities.getElement('security').getValue()
            vals.append({'name': sec, 'cur_mkt_val':None, 'px_last':None})

    return vals

# if __name__ == "__main__":
#     # print("SimpleHistoryExample")
#     try:
#         vals = main()
#
#         print(vals)
#     except KeyboardInterrupt:
#         print("Ctrl+C pressed. Stopping...")
#

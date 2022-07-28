
using System;
using System.Threading;
using System.Runtime.InteropServices;
using Microsoft.FlightSimulator.SimConnect;

namespace MSFSFlightDataDemo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Creating an instance of the SimVarRequester class to get things started
            SimVarRequester svRequester = new SimVarRequester();
        }

    }

    // The SimVarRequester contains all the functionality for opening communication to MSFS, requesting, receiving, and outputting simulation data
    public class SimVarRequester
    {

        private SimConnect simConnect = null;

        // These enumerations dont accomplish much, but their use is a requirement of the SimConnect SDK
        enum DATA_REQUESTS
        {
            DataRequest
        };
        enum DEFINITIONS
        {
            SimPlaneDataStructure
        };

        const int WM_USER_SIMCONNECT = 0x0402; // This is some SDK defined constant necessary for connection
        const int CX_RETRY_SECONDS = 5; // How long to wait to retry connection attempt after failure
        const double DATA_POLLING_SECONDS = 0.5; // How often to request updated sim values freom MSFS

        // This structure is used to instruct SimConnect how to package the data being sent in response to our requests
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        public struct SimPlaneDataStructure
        {
            public double plane_pitch;
            public double plane_bank;
            public double acceleration_z;
        }

        // This constructor initializes the connection to MSFS, configures the requests, and uses a System.Threading.Timer to poll MSFS for updated simulation data
        public SimVarRequester()
        {

            Console.WriteLine("Attempting to connect to MSFS");

            // Attempt connection to MSFS
            // If connection is not successful, then pause and retry until success
            while (simConnect == null)
            {
                try
                {
                    simConnect = new SimConnect("MSFSFlightDataDemo Connection", IntPtr.Zero, WM_USER_SIMCONNECT, null, 0);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Unable to establish connection to MSFS");
                    Console.WriteLine("Error -> [{0}]", e.Message);
                    Console.WriteLine("Retrying in {0} second{1}", CX_RETRY_SECONDS, CX_RETRY_SECONDS == 1 ? "" : "s");
                    Thread.Sleep(CX_RETRY_SECONDS * 1000);
                }
            }

            // Setting up event handlers for open, quit, and excetpions
            simConnect.OnRecvOpen += new SimConnect.RecvOpenEventHandler(SimConnect_OnRecvOpen);
            simConnect.OnRecvQuit += new SimConnect.RecvQuitEventHandler(SimConnect_OnRecvQuit);
            simConnect.OnRecvException += new SimConnect.RecvExceptionEventHandler(SimConnect_OnRecvException);

            // Create a profile for the type of simulations variables we are requesting 
            simConnect.AddToDataDefinition(DEFINITIONS.SimPlaneDataStructure, "Plane Pitch Degrees", "radians", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
            simConnect.AddToDataDefinition(DEFINITIONS.SimPlaneDataStructure, "Plane Bank Degrees", "radians", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
            simConnect.AddToDataDefinition(DEFINITIONS.SimPlaneDataStructure, "Acceleration Body Z", "feet per second squared", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);

            // Instruct SimConnect to use our SimPlaneDataStructure object to package our response data
            simConnect.RegisterDataDefineStruct<SimPlaneDataStructure>(DEFINITIONS.SimPlaneDataStructure);

            // Set up event handler for receiving responses to our requests
            simConnect.OnRecvSimobjectDataBytype += new SimConnect.RecvSimobjectDataBytypeEventHandler(SimConnect_OnRecvSimobjectDataBytype);

            // Create a timer that will issue simulation data requests to SimConnect at a defined interval
            Timer timer = new Timer(this.OnTimerTick, null, 0, (long)(DATA_POLLING_SECONDS * 1000));

            // Create an infinite loop to keep the console application open while the Timer is doing its job
            while (true)
            {
                Thread.Sleep(1000);
            }

        }

        void OnTimerTick(object state)
        {
            try
            {
                // Issue our request for the latest simulation data
                simConnect.RequestDataOnSimObjectType(DATA_REQUESTS.DataRequest, DEFINITIONS.SimPlaneDataStructure, 0, SIMCONNECT_SIMOBJECT_TYPE.USER);
                // Indicate that we are ready to receive a response to that request (silly that this is necessary)
                simConnect.ReceiveMessage();
            }
            catch (Exception e)
            {
                Console.WriteLine("Unable to request data from MSFS");
                Console.WriteLine("Error -> [{0}]", e.Message);
            }
        }

        // Handle connection open events
        void SimConnect_OnRecvOpen(SimConnect sender, SIMCONNECT_RECV_OPEN data)
        {
            Console.WriteLine("Connection to MSFS established");
        }

        // Handle connection quit events
        void SimConnect_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data)
        {
            if (simConnect != null)
            {
                simConnect.Dispose();
                simConnect = null;
                Console.WriteLine("Connection to MSFS has been closed");
            }
        }

        // Handle connection exception events
        void SimConnect_OnRecvException(SimConnect sender, SIMCONNECT_RECV_EXCEPTION data)
        {
            Console.WriteLine("Connection to MSFS has encountered an exception");
            Console.WriteLine("Error -> [{0}]", data.dwException);
        }

        // Handle request response events
        void SimConnect_OnRecvSimobjectDataBytype(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE data)
        {
            // Isolate the requested simulation data from the response object and output the information
            SimPlaneDataStructure currentSimData = (SimPlaneDataStructure)data.dwData[0];
            Console.Write("\rPitch (rad):\t{0:N6}\tBank (rad):\t{1:N6}\tAccZ (f/s^2):\t{2:N6}\t\t", currentSimData.plane_pitch, currentSimData.plane_bank, currentSimData.acceleration_z);
        }
    }
}

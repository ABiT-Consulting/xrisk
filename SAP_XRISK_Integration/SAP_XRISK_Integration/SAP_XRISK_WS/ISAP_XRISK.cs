using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.ServiceModel.Web;
using System.Text;
using System.Web;
using System.Xml.Linq;

namespace SAP_XRISK_WS
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface ISAP_XRISK
    {

        [OperationContract]
        void SaveConnection(string Connection,out string msg);

        [OperationContract]
        string GetConnection();

        [OperationContract]
        int GetXmlandVersion(int version,string objectCode,string criteria,out string xmlobj);
        [OperationContract]
        bool addChangeTracks(string objcode, string start, string end,string trtype="A");

        [OperationContract]
        string getSchema(int VersionNumber);

        [OperationContract]
        bool XRiskToSap(string xml,out string msg, out string KeyValue);
        // TODO: Add your service operations here
    }


    // Use a data contract as illustrated in the sample below to add composite types to service operations.
    //[DataContract]
    //public class CompositeType
    //{
    //    bool boolValue = true;
    //    string stringValue = "Hello ";

    //    [DataMember]
    //    public bool BoolValue
    //    {
    //        get { return boolValue; }
    //        set { boolValue = value; }
    //    }

    //    [DataMember]
    //    public string StringValue
    //    {
    //        get { return stringValue; }
    //        set { stringValue = value; }
    //    }
    //}
}

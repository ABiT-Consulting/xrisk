using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace ComModel
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IComModelService" in both code and config file together.
    [ServiceContract]
    public interface IComModelService
    {
        /// <summary>
        /// This function will return Session ID as per UserName and Password. This Session ID will be used for Transections
        /// </summary>
        /// <param name="username">SAP user name</param>
        /// <param name="password">SAP Password</param>
        /// <param name="Message">output message as returned by authentication function</param>
        /// <returns></returns>
        /// 
        [OperationContract]
        void Login(string username, string password, out String Message, out String SessionID);
        [OperationContract]
        JSON_OBJ GetObjectByKey(String Key, object ObjType, string SessionID);

        [OperationContract]
        Response AddUpdateObject(JSON_OBJ Object, string SessionID);

        [OperationContract]
        void LogOut(String SessionID);
            [OperationContract]
        void GetObjectKeys(object ObjType, string SessionID , out List<object > List);
        /// <param name="ObjKeys">List of all Keys associated to the object type</param>
        //[OperationContract]
        //void GetObjectKeysList(object ObjType, String SessionID, out List<object> ObjKeys);
        // TODO: Add your service operations here
    }

    // Use a data contract as illustrated in the sample below to add composite types to service operations.
    // You can add XSD files into the project. After building the project, you can directly use the data types defined there, with the namespace "ComModel.ContractType".
    [DataContract]
    public class JSON_OBJ
    {
        string jsonObject = "Hello ";
        string ObjectType = "";
        [DataMember]
        public string OBJECT_Type
        {
            get { return ObjectType; }
            set { ObjectType = value; }
        }
        /// <summary>
        /// JSON DATA Packet
        /// </summary>
        [DataMember]
        public string JSON_OBJECT
        {
            get { return jsonObject; }
            set { jsonObject = value; }
        }
        string _Message = "";
        [DataMember]
        public string Message
        {
            get { return _Message; }
            set { _Message = value; }
        }


    }
    [DataContract]
    public class Response
    {
        
        string stringValue = "Hello ";

        /// <summary>
        /// JSON DATA Packet
        /// 1. 
        /// </summary>
        [DataMember]
        public string JSON_Response
        {
            get { return stringValue; }
            set { stringValue = value; }
        }
    }
}

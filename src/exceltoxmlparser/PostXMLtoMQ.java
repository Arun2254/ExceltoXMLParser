package exceltoxmlparser;

import java.io.IOException;
import java.util.Hashtable;

//import com.ibm.mq.MQC;
//import com.ibm.mq.MQException;
//import com.ibm.mq.MQMessage;
//import com.ibm.mq.MQPutMessageOptions;
//import com.ibm.mq.MQQueueManager;
//import com.ibm.mq.constants.MQConstants;
//import com.ibm.mq.MQQueue;
//import com.ibm.msg.client.wmq.WMQConstants;
//
//import com.ibm.mq.MQAsyncStatus;
//import com.ibm.mq.MQException;
//import com.ibm.mq.MQMessage;
//import com.ibm.mq.MQQueueManager;
//import com.ibm.mq.constants.MQConstants;


import com.ibm.mq.MQAsyncStatus;
import com.ibm.mq.MQException;
import com.ibm.mq.MQMessage;
import com.ibm.mq.MQPutMessageOptions;
import com.ibm.mq.MQQueue;
import com.ibm.mq.MQQueueManager;
import com.ibm.mq.constants.MQConstants;


public class PostXMLtoMQ {
	
	public void postmsgtoMQ(String pomessage) throws IOException{
//		Hashtable<String, Object> props = new Hashtable<String, Object>();
//        props.put(MQConstants.CHANNEL_PROPERTY, "WMM.SVRCONN");
//        props.put(MQConstants.PORT_PROPERTY,  30011);
//        props.put(MQConstants.HOST_NAME_PROPERTY, "tgphxwmmmq001.phx.gapinc.dev");
//
        String qManager = "WMM01I";
        String queueName = "QL.04.TEST.POET";
        MQQueueManager qMgr = null;
        MQPutMessageOptions pmo =null;
        
        try{

            //Create a Hashtable with required properties
            Hashtable<String, Object> properties = new Hashtable<String, Object>();
            properties.put(MQConstants.CHANNEL_PROPERTY, "WMM.SVRCONN");
	        properties.put(MQConstants.PORT_PROPERTY,  30011);
	        properties.put(MQConstants.HOST_NAME_PROPERTY, "tgphxwmmmq001.phx.gapinc.dev"); 

            //Create a instance of qManager
            qMgr = new MQQueueManager(qManager, properties);
            int openOptions = MQConstants.MQOO_OUTPUT | MQConstants.MQOO_INPUT_AS_Q_DEF;
            //Connect to the Queue
            MQQueue queue = qMgr.accessQueue(queueName, openOptions);
            pmo = new MQPutMessageOptions();
            pmo.options = MQConstants.MQPMO_ASYNC_RESPONSE;
            

            //Creating the mqmessage
            MQMessage mqMsg = new MQMessage();
            mqMsg.format = MQConstants.MQFMT_STRING;
            mqMsg.writeString(pomessage);
            System.out.println("Message Posted Sucessfully into "+queueName);
            queue.put(mqMsg,pmo);
            queue.close();
            qMgr.disconnect();
            
         }catch(MQException mqEx){
            mqEx.printStackTrace();
         }
	}

}

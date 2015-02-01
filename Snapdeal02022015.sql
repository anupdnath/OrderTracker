CREATE DATABASE  IF NOT EXISTS `ordertracker` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `ordertracker`;
-- MySQL dump 10.13  Distrib 5.6.17, for Win32 (x86)
--
-- Host: 127.0.0.1    Database: ordertracker
-- ------------------------------------------------------
-- Server version	5.6.21-log

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `orderallamount`
--

DROP TABLE IF EXISTS `orderallamount`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orderallamount` (
  `Suborderid` varchar(45) NOT NULL,
  `COD_NON_COD_Credit` decimal(18,2) DEFAULT '0.00',
  `COD_NON_COD_Debit` decimal(18,2) DEFAULT '0.00',
  `Incentive` decimal(18,2) DEFAULT '0.00',
  `Disincentive` decimal(18,2) DEFAULT '0.00',
  `COD_NCOD_Wrong_faulty_Debit` decimal(18,2) DEFAULT '0.00',
  `Stock_Out_Commission` decimal(18,2) DEFAULT '0.00',
  `Courier_lost_vendor` decimal(18,2) DEFAULT '0.00',
  `COD_Non_COD_Frgt_post_ship` decimal(18,2) DEFAULT '0.00',
  `RTO_Conflict` decimal(18,2) DEFAULT '0.00',
  PRIMARY KEY (`Suborderid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orderallamount`
--

LOCK TABLES `orderallamount` WRITE;
/*!40000 ALTER TABLE `orderallamount` DISABLE KEYS */;
/*!40000 ALTER TABLE `orderallamount` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `orderdetails`
--

DROP TABLE IF EXISTS `orderdetails`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orderdetails` (
  `SubOrderID` varchar(45) NOT NULL,
  `Status` varchar(45) DEFAULT NULL,
  `Remark` varchar(128) DEFAULT NULL,
  `UpdatedDate` datetime DEFAULT NULL,
  `Amount` decimal(18,2) DEFAULT NULL,
  `CreationDate` datetime DEFAULT NULL,
  PRIMARY KEY (`SubOrderID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orderdetails`
--

LOCK TABLES `orderdetails` WRITE;
/*!40000 ALTER TABLE `orderdetails` DISABLE KEYS */;
/*!40000 ALTER TABLE `orderdetails` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `orderhos`
--

DROP TABLE IF EXISTS `orderhos`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orderhos` (
  `suborderid` varchar(45) NOT NULL,
  `Sku` varchar(100) DEFAULT NULL,
  `Supc` varchar(100) DEFAULT NULL,
  `AWB` varchar(100) DEFAULT NULL,
  `Ref` varchar(100) DEFAULT NULL,
  `CreationDate` datetime DEFAULT NULL,
  `hosno` varchar(45) DEFAULT NULL,
  `hosdate` varchar(45) DEFAULT NULL,
  `hsdate` datetime DEFAULT NULL,
  `weight` varchar(45) DEFAULT NULL,
  `mobile` varchar(45) DEFAULT NULL,
  `RecDetails` varchar(500) DEFAULT NULL,
  `Manifestid` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`suborderid`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orderhos`
--

LOCK TABLES `orderhos` WRITE;
/*!40000 ALTER TABLE `orderhos` DISABLE KEYS */;
/*!40000 ALTER TABLE `orderhos` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `orderpacked`
--

DROP TABLE IF EXISTS `orderpacked`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orderpacked` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `suborderid` varchar(45) NOT NULL DEFAULT '',
  `Category` varchar(100) DEFAULT NULL,
  `Courier` varchar(100) DEFAULT NULL,
  `Product` varchar(100) DEFAULT NULL,
  `Reference_Code` varchar(100) DEFAULT NULL,
  `SKU_Code` varchar(100) DEFAULT NULL,
  `AWB_Number` varchar(100) DEFAULT NULL,
  `Order_Verified_Date` datetime DEFAULT NULL,
  `Order_Created_Date` datetime DEFAULT NULL,
  `Customer_Name` varchar(100) DEFAULT NULL,
  `Shipping_Name` varchar(100) DEFAULT NULL,
  `City` varchar(100) DEFAULT NULL,
  `State` varchar(100) DEFAULT NULL,
  `PIN_Code` varchar(100) DEFAULT NULL,
  `Selling_Price_Per_Item` varchar(100) DEFAULT NULL,
  `IMEI_SERIAL` varchar(100) DEFAULT NULL,
  `PromisedShipDate` datetime DEFAULT NULL,
  `MRP` varchar(100) DEFAULT NULL,
  `InvoiceCode` varchar(100) DEFAULT NULL,
  `CreationDate` datetime DEFAULT NULL,
  PRIMARY KEY (`id`,`suborderid`)
) ENGINE=InnoDB AUTO_INCREMENT=154 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orderpacked`
--

LOCK TABLES `orderpacked` WRITE;
/*!40000 ALTER TABLE `orderpacked` DISABLE KEYS */;
/*!40000 ALTER TABLE `orderpacked` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `orderpayment`
--

DROP TABLE IF EXISTS `orderpayment`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orderpayment` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `suborderid` varchar(45) DEFAULT NULL,
  `customername` varchar(100) DEFAULT NULL,
  `web_sale_price` varchar(100) DEFAULT NULL,
  `cash_back` varchar(100) DEFAULT NULL,
  `emi` varchar(100) DEFAULT NULL,
  `merchant_cut` varchar(100) DEFAULT NULL,
  `fulfillment_fee` varchar(100) DEFAULT NULL,
  `fulfillment_fee_waiver` varchar(100) DEFAULT NULL,
  `marketing_fee` varchar(100) DEFAULT NULL,
  `payment_collection_fee` varchar(100) DEFAULT NULL,
  `courier_fee` varchar(100) DEFAULT NULL,
  `comm` varchar(100) DEFAULT NULL,
  `sku_code` varchar(100) DEFAULT NULL,
  `shipped_return_date` varchar(100) DEFAULT NULL,
  `delivered_date` varchar(100) DEFAULT NULL,
  `shipping_method_code` varchar(100) DEFAULT NULL,
  `courier` varchar(100) DEFAULT NULL,
  `awb_number` varchar(100) DEFAULT NULL,
  `amount` decimal(18,2) DEFAULT NULL,
  `other_applications` varchar(100) DEFAULT NULL,
  `creationdate` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orderpayment`
--

LOCK TABLES `orderpayment` WRITE;
/*!40000 ALTER TABLE `orderpayment` DISABLE KEYS */;
/*!40000 ALTER TABLE `orderpayment` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `orderstatus`
--

DROP TABLE IF EXISTS `orderstatus`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `orderstatus` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `statusname` varchar(45) DEFAULT NULL,
  `priority` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `orderstatus`
--

LOCK TABLES `orderstatus` WRITE;
/*!40000 ALTER TABLE `orderstatus` DISABLE KEYS */;
INSERT INTO `orderstatus` (`id`, `statusname`, `priority`) VALUES (1,'Packed',1),(2,'Shipped',2),(3,'Payment Received',3),(4,'Penalty',4),(5,'Returned',5),(6,'Penalty Approved',6),(7,'All',0),(8,'RTO Recieved',7),(9,'Customer Complaint Acknowledged',8);
/*!40000 ALTER TABLE `orderstatus` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `ordertransection`
--

DROP TABLE IF EXISTS `ordertransection`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `ordertransection` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `suborderid` varchar(45) NOT NULL,
  `remark` varchar(128) DEFAULT NULL,
  `creationDate` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=4336 DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `ordertransection`
--

LOCK TABLES `ordertransection` WRITE;
/*!40000 ALTER TABLE `ordertransection` DISABLE KEYS */;
/*!40000 ALTER TABLE `ordertransection` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `paymentmap`
--

DROP TABLE IF EXISTS `paymentmap`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `paymentmap` (
  `ActualField` varchar(100) NOT NULL,
  `Tablefield` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`ActualField`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paymentmap`
--

LOCK TABLES `paymentmap` WRITE;
/*!40000 ALTER TABLE `paymentmap` DISABLE KEYS */;
/*!40000 ALTER TABLE `paymentmap` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `paymentstatus`
--

DROP TABLE IF EXISTS `paymentstatus`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `paymentstatus` (
  `code` varchar(45) DEFAULT NULL,
  `status` varchar(45) DEFAULT NULL,
  `remark` varchar(45) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `paymentstatus`
--

LOCK TABLES `paymentstatus` WRITE;
/*!40000 ALTER TABLE `paymentstatus` DISABLE KEYS */;
INSERT INTO `paymentstatus` (`code`, `status`, `remark`) VALUES ('COD','Payment Received',NULL),('COD','Sales Return','Missing Credit'),('NCOD','Sales Return','Missing Credit'),('NCOD','Payment Received',''),('COD','Sales Return','Subject to be Approved'),('NCOD','Sales Return','Subject to be Approved'),('Incentive','','Payment Not Received'),('Disincentive','','Payment Not Received'),('COD Wrong faulty','Customer Complaint','Penalty Charged'),('NCOD Wrong faulty','Customer Complaint','Penalty Charged'),('Stock Out Commission','Stock out commission',''),('Courier Lost Vendor','Payment received','Courier Lost Vendor'),('COD Frgt Post Ship','Sales return','Penalty Charged'),('NCOD Frgt Post Ship','Sales return','Penalty Charged'),('RTO Conflict','Payment received','RTO Conflict');
/*!40000 ALTER TABLE `paymentstatus` ENABLE KEYS */;
UNLOCK TABLES;

--
-- Table structure for table `usermaster`
--

DROP TABLE IF EXISTS `usermaster`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
/*!40101 SET character_set_client = utf8 */;
CREATE TABLE `usermaster` (
  `username` varchar(50) NOT NULL,
  `Pass` varchar(45) DEFAULT NULL,
  `UserType` varchar(45) DEFAULT NULL,
  `CreatedOn` datetime DEFAULT NULL,
  `UpdatedOn` datetime DEFAULT NULL,
  `Active` bit(1) DEFAULT NULL,
  PRIMARY KEY (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `usermaster`
--

LOCK TABLES `usermaster` WRITE;
/*!40000 ALTER TABLE `usermaster` DISABLE KEYS */;
INSERT INTO `usermaster` (`username`, `Pass`, `UserType`, `CreatedOn`, `UpdatedOn`, `Active`) VALUES ('admin','admin','admin',NULL,NULL,''),('test','121212','USER','2014-12-03 15:38:39','2014-12-03 16:18:05','');
/*!40000 ALTER TABLE `usermaster` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2015-02-02  0:10:03

CREATE DATABASE  IF NOT EXISTS `mempoi` /*!40100 DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci */;
USE `mempoi`;
-- MySQL dump 10.13  Distrib 8.0.15, for macos10.14 (x86_64)
--
-- Host: 127.0.0.1    Database: mempoi
-- ------------------------------------------------------
-- Server version	8.0.15

/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
 SET NAMES utf8 ;
/*!40103 SET @OLD_TIME_ZONE=@@TIME_ZONE */;
/*!40103 SET TIME_ZONE='+00:00' */;
/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;
/*!40111 SET @OLD_SQL_NOTES=@@SQL_NOTES, SQL_NOTES=0 */;

--
-- Table structure for table `export_test`
--

DROP TABLE IF EXISTS `export_test`;
/*!40101 SET @saved_cs_client     = @@character_set_client */;
 SET character_set_client = utf8mb4 ;
CREATE TABLE `export_test` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT,
  `creation_date` date NOT NULL,
  `dateTime` datetime DEFAULT NULL,
  `timeStamp` timestamp NULL DEFAULT NULL,
  `name` varchar(100) COLLATE utf8mb4_general_ci NOT NULL,
  `valid` tinyint(1) DEFAULT NULL,
  `usefulChar` char(1) COLLATE utf8mb4_general_ci DEFAULT NULL,
  `decimalOne` decimal(10,0) DEFAULT NULL,
  `bitTwo` bit(1) DEFAULT NULL,
  `doublone` double DEFAULT NULL,
  `floattone` float DEFAULT NULL,
  `interao` int(11) DEFAULT NULL,
  `mediano` mediumint(9) DEFAULT NULL,
  `attempato` time DEFAULT NULL,
  `interuccio` tinyint(4) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;
/*!40101 SET character_set_client = @saved_cs_client */;

--
-- Dumping data for table `export_test`
--

LOCK TABLES `export_test` WRITE;
/*!40000 ALTER TABLE `export_test` DISABLE KEYS */;
INSERT INTO `export_test` VALUES (1,'2019-01-04','2019-01-04 15:30:14','2018-01-02 15:12:50','tancredi',1,'T',10,_binary '\0',12392112923,45.1,2500,167,'03:10:34',12),(2,'2019-01-04','2019-01-04 15:30:14','2018-01-02 15:12:50','tancredi',0,'T',16,_binary '\0',12392112923,45.1,2500,167,'03:10:34',12),(3,'2019-01-04','2019-01-04 15:30:14','2018-01-02 15:12:50','tancredi',1,'T',110,_binary '\0',12392112923,45.1,2500,167,'03:10:34',12),(4,'2019-01-04','2019-01-04 15:30:14','2018-01-02 15:12:50','tancredi',0,'T',130,_binary '\0',12392112923,45.1,2500,167,'03:10:34',12),(5,'2019-01-04','2019-01-04 15:30:14','2018-01-02 15:12:50','tancredi',1,'T',1055,_binary '\0',12392112923,45.1,2500,167,'03:10:34',12);
/*!40000 ALTER TABLE `export_test` ENABLE KEYS */;
UNLOCK TABLES;
/*!40103 SET TIME_ZONE=@OLD_TIME_ZONE */;

/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40111 SET SQL_NOTES=@OLD_SQL_NOTES */;

-- Dump completed on 2019-04-16 10:06:29

-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	5.1.59-community


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


--
-- Create schema dbinventory
--

CREATE DATABASE IF NOT EXISTS dbinventory;
USE dbinventory;

--
-- Temporary table structure for view `view_ending_balance`
--
DROP TABLE IF EXISTS `view_ending_balance`;
DROP VIEW IF EXISTS `view_ending_balance`;
CREATE TABLE `view_ending_balance` (
  `item_code` varchar(45),
  `item_qty` double(10,2)
);

--
-- Definition of table `account_receivable`
--

DROP TABLE IF EXISTS `account_receivable`;
CREATE TABLE `account_receivable` (
  `sales_order_no` varchar(45) NOT NULL DEFAULT '0',
  `remarks` varchar(45) DEFAULT NULL,
  `date` datetime DEFAULT NULL,
  PRIMARY KEY (`sales_order_no`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `account_receivable`
--

/*!40000 ALTER TABLE `account_receivable` DISABLE KEYS */;
INSERT INTO `account_receivable` (`sales_order_no`,`remarks`,`date`) VALUES 
 ('SO-0000036','','2011-11-05 23:13:29'),
 ('SO-0000040','','2011-11-05 23:39:57'),
 ('SO-0000042','','2011-11-05 23:45:26'),
 ('SO-0000052','','2011-11-06 02:19:44'),
 ('SO-0000055','','2011-11-13 15:31:55'),
 ('SO-0000056','','2011-11-13 15:53:03'),
 ('SO-0000057','','2011-11-13 15:56:57'),
 ('SO-000008','',NULL);
/*!40000 ALTER TABLE `account_receivable` ENABLE KEYS */;


--
-- Definition of table `account_recievable_cart`
--

DROP TABLE IF EXISTS `account_recievable_cart`;
CREATE TABLE `account_recievable_cart` (
  `acount_recievable_cart_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `item_code` varchar(45) NOT NULL,
  `qty` int(10) unsigned NOT NULL,
  `customer_type` varchar(45) NOT NULL,
  `acount_recievable_cart_date` datetime NOT NULL,
  `price` double(2,2) NOT NULL,
  `total_price` double(2,2) NOT NULL,
  `discount_amount` double(2,2) NOT NULL,
  PRIMARY KEY (`acount_recievable_cart_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `account_recievable_cart`
--

/*!40000 ALTER TABLE `account_recievable_cart` DISABLE KEYS */;
/*!40000 ALTER TABLE `account_recievable_cart` ENABLE KEYS */;


--
-- Definition of table `account_recievable_payments`
--

DROP TABLE IF EXISTS `account_recievable_payments`;
CREATE TABLE `account_recievable_payments` (
  `account_recievable_id` int(10) unsigned NOT NULL DEFAULT '0',
  `payment_id` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`account_recievable_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `account_recievable_payments`
--

/*!40000 ALTER TABLE `account_recievable_payments` DISABLE KEYS */;
/*!40000 ALTER TABLE `account_recievable_payments` ENABLE KEYS */;


--
-- Definition of table `account_recievable_to_account_recievable_cart`
--

DROP TABLE IF EXISTS `account_recievable_to_account_recievable_cart`;
CREATE TABLE `account_recievable_to_account_recievable_cart` (
  `account_recievable_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `id` int(10) unsigned NOT NULL,
  PRIMARY KEY (`account_recievable_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `account_recievable_to_account_recievable_cart`
--

/*!40000 ALTER TABLE `account_recievable_to_account_recievable_cart` DISABLE KEYS */;
/*!40000 ALTER TABLE `account_recievable_to_account_recievable_cart` ENABLE KEYS */;


--
-- Definition of table `agent`
--

DROP TABLE IF EXISTS `agent`;
CREATE TABLE `agent` (
  `agent_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Name` varchar(45) NOT NULL,
  `Mobile` varchar(45) DEFAULT NULL,
  `address` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`agent_id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `agent`
--

/*!40000 ALTER TABLE `agent` DISABLE KEYS */;
INSERT INTO `agent` (`agent_id`,`Name`,`Mobile`,`address`) VALUES 
 (1,'Nic Lomocso','09129720502','Bilar'),
 (2,'Ronilo Lopiseros','09183395029','ad'),
 (3,'Dario Medina','09286321347','adr'),
 (4,'Robert Idulza','09125387404','Bilar'),
 (5,'Jerson Samuya','09494063344','San Isidro');
/*!40000 ALTER TABLE `agent` ENABLE KEYS */;


--
-- Definition of table `agent_customers`
--

DROP TABLE IF EXISTS `agent_customers`;
CREATE TABLE `agent_customers` (
  `agent_id` int(10) unsigned DEFAULT NULL,
  `customers_id` int(10) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `agent_customers`
--

/*!40000 ALTER TABLE `agent_customers` DISABLE KEYS */;
INSERT INTO `agent_customers` (`agent_id`,`customers_id`) VALUES 
 (1,24),
 (3,8),
 (1,9),
 (5,10);
/*!40000 ALTER TABLE `agent_customers` ENABLE KEYS */;


--
-- Definition of table `cart`
--

DROP TABLE IF EXISTS `cart`;
CREATE TABLE `cart` (
  `cart_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `item_code` varchar(45) NOT NULL,
  `qty` int(10) unsigned NOT NULL,
  `customer_type` varchar(45) NOT NULL,
  `cart_date` datetime NOT NULL,
  `price` double(2,2) NOT NULL,
  `total_price` double(2,2) NOT NULL,
  `discount` varchar(45) NOT NULL,
  `discount_amount` double(2,2) NOT NULL,
  PRIMARY KEY (`cart_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cart`
--

/*!40000 ALTER TABLE `cart` DISABLE KEYS */;
/*!40000 ALTER TABLE `cart` ENABLE KEYS */;


--
-- Definition of table `cod`
--

DROP TABLE IF EXISTS `cod`;
CREATE TABLE `cod` (
  `sales_order_no` varchar(45) NOT NULL DEFAULT '0',
  `remarks` varchar(45) DEFAULT NULL,
  `date` datetime DEFAULT NULL,
  PRIMARY KEY (`sales_order_no`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cod`
--

/*!40000 ALTER TABLE `cod` DISABLE KEYS */;
INSERT INTO `cod` (`sales_order_no`,`remarks`,`date`) VALUES 
 ('SO-0000010','',NULL),
 ('SO-0000011','','2011-11-05 11:01:48'),
 ('SO-0000012','','2011-11-05 11:37:53'),
 ('SO-0000013','','2011-11-05 11:38:32'),
 ('SO-0000014','','2011-11-05 11:40:20'),
 ('SO-0000015','','2011-11-05 11:42:00'),
 ('SO-0000016','','2011-11-05 11:43:59'),
 ('SO-0000017','','2011-11-05 11:44:32'),
 ('SO-0000018','','2011-11-05 11:45:57'),
 ('SO-0000019','','2011-11-05 11:47:52'),
 ('SO-0000020','','2011-11-05 11:56:20'),
 ('SO-0000021','','2011-11-05 12:16:31'),
 ('SO-0000022','','2011-11-05 12:18:10'),
 ('SO-0000023','','2011-11-05 12:19:04'),
 ('SO-0000024','','2011-11-05 13:17:55'),
 ('SO-0000025','','2011-11-05 13:19:15'),
 ('SO-0000026','','2011-11-05 13:19:51'),
 ('SO-0000027','','2011-11-05 20:19:47'),
 ('SO-0000028','','2011-11-05 20:38:01'),
 ('SO-0000029','','2011-11-05 20:43:31'),
 ('SO-000003','',NULL),
 ('SO-0000030','','2011-11-05 21:10:45'),
 ('SO-0000031','','2011-11-05 21:11:36'),
 ('SO-0000032','','2011-11-05 21:12:33'),
 ('SO-0000033','','2011-11-05 21:13:49'),
 ('SO-0000034','','2011-11-05 21:19:53'),
 ('SO-0000035','','2011-11-05 21:58:27'),
 ('SO-0000037','','2011-11-05 23:27:28'),
 ('SO-0000038','','2011-11-05 23:32:08'),
 ('SO-0000039','','2011-11-05 23:34:24'),
 ('SO-0000041','','2011-11-05 23:40:36'),
 ('SO-0000043','','2011-11-05 23:45:50'),
 ('SO-0000044','','2011-11-05 23:47:59'),
 ('SO-0000045','','2011-11-05 23:51:14'),
 ('SO-0000046','','2011-11-05 23:54:51'),
 ('SO-0000047','','2011-11-06 00:24:34'),
 ('SO-0000048','','2011-11-06 00:27:00'),
 ('SO-0000049','','2011-11-06 00:28:30'),
 ('SO-0000050','','2011-11-06 00:29:53'),
 ('SO-0000051','','2011-11-06 00:37:49'),
 ('SO-0000053','','2011-11-12 18:17:21'),
 ('SO-0000054','','2011-11-13 14:52:04'),
 ('SO-000006','',NULL),
 ('SO-000009','',NULL);
/*!40000 ALTER TABLE `cod` ENABLE KEYS */;


--
-- Definition of table `customers`
--

DROP TABLE IF EXISTS `customers`;
CREATE TABLE `customers` (
  `customers_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `customers_name` varchar(45) NOT NULL,
  `customers_add` varchar(45) NOT NULL,
  `customers_number` varchar(45) NOT NULL,
  `dealers_type` varchar(45) NOT NULL,
  PRIMARY KEY (`customers_id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=24 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `customers`
--

/*!40000 ALTER TABLE `customers` DISABLE KEYS */;
INSERT INTO `customers` (`customers_id`,`customers_name`,`customers_add`,`customers_number`,`dealers_type`) VALUES 
 (8,'Alfie Ocba','Tagbilaran','123','dealer'),
 (9,'Alvin Crecencio','Talibon','','consumer'),
 (10,'Andres Arong','Ubay','','dealer'),
 (11,'Anita Cañada','Tubigon','',''),
 (12,'Annie Aparre','Inabanga','',''),
 (13,'Arlene Abres','Cahayag','',''),
 (14,'Aster Galola','tubigon','',''),
 (15,'Baloy Blantucas','tubigon','',''),
 (16,'Bebeth Barrosa','tubigon','',''),
 (17,'Bijie','Carmen','',''),
 (18,'BISU','Calape','',''),
 (19,'Boy Capio','Inabanga','',''),
 (20,'Cabandos','Trinidad','',''),
 (21,'Ben Gablines','Calape','',''),
 (22,'Beneth Resonabe','Calape','',''),
 (23,'Arlene Amasor','Desamparados, Calape, Bohol','','');
/*!40000 ALTER TABLE `customers` ENABLE KEYS */;


--
-- Definition of table `customers_discount`
--

DROP TABLE IF EXISTS `customers_discount`;
CREATE TABLE `customers_discount` (
  `customers_id` int(10) unsigned NOT NULL DEFAULT '0',
  `discount_id` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`customers_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `customers_discount`
--

/*!40000 ALTER TABLE `customers_discount` DISABLE KEYS */;
/*!40000 ALTER TABLE `customers_discount` ENABLE KEYS */;


--
-- Definition of table `discount`
--

DROP TABLE IF EXISTS `discount`;
CREATE TABLE `discount` (
  `discount_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `discount_code` varchar(45) DEFAULT NULL,
  `discount_name` varchar(45) DEFAULT NULL,
  `amount` double(10,2) DEFAULT NULL,
  PRIMARY KEY (`discount_id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `discount`
--

/*!40000 ALTER TABLE `discount` DISABLE KEYS */;
INSERT INTO `discount` (`discount_id`,`discount_code`,`discount_name`,`amount`) VALUES 
 (2,'123','sample',50.00);
/*!40000 ALTER TABLE `discount` ENABLE KEYS */;


--
-- Definition of table `inventory`
--

DROP TABLE IF EXISTS `inventory`;
CREATE TABLE `inventory` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `item_id` int(11) DEFAULT NULL,
  `item_code` varchar(45) DEFAULT NULL,
  `beginning_balance` double(10,2) DEFAULT NULL,
  `ending_balance` double(10,2) DEFAULT NULL,
  `date` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=124 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `inventory`
--

/*!40000 ALTER TABLE `inventory` DISABLE KEYS */;
INSERT INTO `inventory` (`id`,`item_id`,`item_code`,`beginning_balance`,`ending_balance`,`date`) VALUES 
 (1,17,'NLB',0.00,0.00,'2011-10-23 15:48:34'),
 (2,18,'NLP',0.00,0.00,'2011-10-23 15:48:34'),
 (3,19,'NSP',0.00,0.00,'2011-10-23 15:48:34'),
 (4,20,'NGP',0.00,0.00,'2011-10-23 15:48:34'),
 (5,21,'NBP',0.00,0.00,'2011-10-23 15:48:34'),
 (6,22,'PSP',0.00,0.00,'2011-10-23 15:48:34'),
 (7,23,'LSP',0.00,0.00,'2011-10-23 15:48:34'),
 (8,25,'BFC',0.00,0.00,'2011-10-23 15:48:34'),
 (9,26,'BSC',0.00,0.00,'2011-10-23 15:48:34'),
 (10,27,'CBM 25',0.00,0.00,'2011-10-23 15:48:34'),
 (11,28,'CGM',0.00,0.00,'2011-10-23 15:48:34'),
 (16,17,'NLB',0.00,0.00,'2011-10-23 15:49:23'),
 (17,18,'NLP',0.00,0.00,'2011-10-23 15:49:23'),
 (18,19,'NSP',0.00,0.00,'2011-10-23 15:49:23'),
 (19,20,'NGP',0.00,0.00,'2011-10-23 15:49:23'),
 (20,21,'NBP',0.00,0.00,'2011-10-23 15:49:23'),
 (21,22,'PSP',0.00,0.00,'2011-10-23 15:49:23'),
 (22,23,'LSP',0.00,0.00,'2011-10-23 15:49:23'),
 (23,25,'BFC',0.00,0.00,'2011-10-23 15:49:23'),
 (24,26,'BSC',0.00,0.00,'2011-10-23 15:49:23'),
 (25,27,'CBM 25',0.00,0.00,'2011-10-23 15:49:23'),
 (26,28,'CGM',0.00,0.00,'2011-10-23 15:49:23'),
 (31,17,'NLB',0.00,0.00,'2011-10-23 17:03:02'),
 (32,18,'NLP',0.00,0.00,'2011-10-23 17:03:02'),
 (33,19,'NSP',0.00,0.00,'2011-10-23 17:03:02'),
 (34,20,'NGP',0.00,0.00,'2011-10-23 17:03:02'),
 (35,21,'NBP',0.00,0.00,'2011-10-23 17:03:02'),
 (36,22,'PSP',0.00,0.00,'2011-10-23 17:03:02'),
 (37,23,'LSP',0.00,0.00,'2011-10-23 17:03:02'),
 (38,25,'BFC',0.00,0.00,'2011-10-23 17:03:02'),
 (39,26,'BSC',0.00,0.00,'2011-10-23 17:03:02'),
 (40,27,'CBM 25',0.00,0.00,'2011-10-23 17:03:02'),
 (41,28,'CGM',0.00,0.00,'2011-10-23 17:03:02'),
 (46,17,'NLB',0.00,0.00,'2011-10-23 17:12:24'),
 (47,18,'NLP',0.00,0.00,'2011-10-23 17:12:24'),
 (48,19,'NSP',0.00,0.00,'2011-10-23 17:12:24'),
 (49,20,'NGP',0.00,0.00,'2011-10-23 17:12:24'),
 (50,21,'NBP',0.00,0.00,'2011-10-23 17:12:24'),
 (51,22,'PSP',0.00,0.00,'2011-10-23 17:12:24'),
 (52,23,'LSP',0.00,0.00,'2011-10-23 17:12:24'),
 (53,25,'BFC',0.00,0.00,'2011-10-23 17:12:24'),
 (54,26,'BSC',0.00,0.00,'2011-10-23 17:12:24'),
 (55,27,'CBM 25',0.00,0.00,'2011-10-23 17:12:24'),
 (56,28,'CGM',0.00,0.00,'2011-10-23 17:12:24'),
 (57,17,'NLB',0.00,0.00,'2011-10-28 11:32:29'),
 (58,18,'NLP',0.00,0.00,'2011-10-28 11:32:29'),
 (59,19,'NSP',0.00,0.00,'2011-10-28 11:32:29'),
 (60,20,'NGP',0.00,0.00,'2011-10-28 11:32:29'),
 (61,21,'NBP',0.00,0.00,'2011-10-28 11:32:29'),
 (62,22,'PSP',0.00,0.00,'2011-10-28 11:32:29'),
 (63,23,'LSP',0.00,0.00,'2011-10-28 11:32:29'),
 (64,25,'BFC',0.00,0.00,'2011-10-28 11:32:29'),
 (65,26,'BSC',0.00,0.00,'2011-10-28 11:32:29'),
 (66,27,'CBM 25',0.00,0.00,'2011-10-28 11:32:29'),
 (67,28,'CGM',0.00,0.00,'2011-10-28 11:32:29'),
 (68,17,'NLB',15.00,10.00,'2011-11-01 18:17:28'),
 (69,18,'NLP',0.00,0.00,'2011-11-01 18:17:28'),
 (70,19,'NSP',150.00,0.00,'2011-11-01 18:17:28'),
 (71,20,'NGP',50.00,0.00,'2011-11-01 18:17:28'),
 (72,21,'NBP',100.00,0.00,'2011-11-01 18:17:28'),
 (73,22,'PSP',301.00,0.00,'2011-11-01 18:17:28'),
 (74,23,'LSP',1601.00,0.00,'2011-11-01 18:17:28'),
 (75,25,'BFC',0.00,0.00,'2011-11-01 18:17:28'),
 (76,26,'BSC',0.00,0.00,'2011-11-01 18:17:28'),
 (77,27,'CBM 25',300.00,0.00,'2011-11-01 18:17:28'),
 (78,28,'CGM',100.00,0.00,'2011-11-01 18:17:28'),
 (79,17,'NLB',96.00,96.00,'2011-11-12 18:00:15'),
 (80,18,'NLP',98.00,98.00,'2011-11-12 18:00:15'),
 (81,19,'NSP',98.00,98.00,'2011-11-12 18:00:15'),
 (82,20,'NGP',100.00,100.00,'2011-11-12 18:00:15'),
 (83,21,'NBP',99.00,99.00,'2011-11-12 18:00:15'),
 (84,22,'PSP',100.00,100.00,'2011-11-12 18:00:15'),
 (85,23,'LSP',99.00,99.00,'2011-11-12 18:00:15'),
 (86,25,'BFC',100.00,100.00,'2011-11-12 18:00:15'),
 (87,26,'BSC',100.00,100.00,'2011-11-12 18:00:15'),
 (88,27,'CBM 25',98.00,98.00,'2011-11-12 18:00:15'),
 (89,28,'CGM',100.00,100.00,'2011-11-12 18:00:15'),
 (90,29,'item2',10.00,10.00,'2011-11-12 18:00:15'),
 (91,30,'item3',40.00,40.00,'2011-11-12 18:00:15'),
 (92,31,'item3',40.00,40.00,'2011-11-12 18:00:15'),
 (93,32,'item5',10.00,10.00,'2011-11-12 18:00:15'),
 (94,17,'NLB',96.00,96.00,'2011-11-12 18:13:36'),
 (95,18,'NLP',98.00,98.00,'2011-11-12 18:13:36'),
 (96,19,'NSP',98.00,98.00,'2011-11-12 18:13:36'),
 (97,20,'NGP',100.00,100.00,'2011-11-12 18:13:36'),
 (98,21,'NBP',99.00,99.00,'2011-11-12 18:13:36'),
 (99,22,'PSP',100.00,100.00,'2011-11-12 18:13:36'),
 (100,23,'LSP',99.00,99.00,'2011-11-12 18:13:36'),
 (101,25,'BFC',100.00,100.00,'2011-11-12 18:13:36'),
 (102,26,'BSC',100.00,100.00,'2011-11-12 18:13:36'),
 (103,27,'CBM 25',98.00,98.00,'2011-11-12 18:13:36'),
 (104,28,'CGM',100.00,100.00,'2011-11-12 18:13:36'),
 (105,29,'item2',10.00,10.00,'2011-11-12 18:13:36'),
 (106,30,'item3',40.00,40.00,'2011-11-12 18:13:36'),
 (107,31,'item3',40.00,40.00,'2011-11-12 18:13:36'),
 (108,32,'item5',10.00,10.00,'2011-11-12 18:13:36'),
 (109,17,'NLB',96.00,96.00,'2011-11-12 18:14:15'),
 (110,18,'NLP',98.00,98.00,'2011-11-12 18:14:15'),
 (111,19,'NSP',98.00,98.00,'2011-11-12 18:14:15'),
 (112,20,'NGP',100.00,100.00,'2011-11-12 18:14:15'),
 (113,21,'NBP',99.00,99.00,'2011-11-12 18:14:15'),
 (114,22,'PSP',100.00,100.00,'2011-11-12 18:14:15'),
 (115,23,'LSP',99.00,99.00,'2011-11-12 18:14:15'),
 (116,25,'BFC',100.00,100.00,'2011-11-12 18:14:15'),
 (117,26,'BSC',100.00,100.00,'2011-11-12 18:14:15'),
 (118,27,'CBM 25',98.00,98.00,'2011-11-12 18:14:15'),
 (119,28,'CGM',100.00,100.00,'2011-11-12 18:14:15'),
 (120,29,'item2',10.00,10.00,'2011-11-12 18:14:15'),
 (121,30,'item3',40.00,40.00,'2011-11-12 18:14:15'),
 (122,31,'item3',40.00,40.00,'2011-11-12 18:14:15'),
 (123,32,'item5',10.00,10.00,'2011-11-12 18:14:15');
/*!40000 ALTER TABLE `inventory` ENABLE KEYS */;


--
-- Definition of table `item_category`
--

DROP TABLE IF EXISTS `item_category`;
CREATE TABLE `item_category` (
  `item_code` varchar(45) DEFAULT NULL,
  `category` varchar(45) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `item_category`
--

/*!40000 ALTER TABLE `item_category` DISABLE KEYS */;
INSERT INTO `item_category` (`item_code`,`category`) VALUES 
 ('NLB','feeds'),
 ('CGM',''),
 ('item2','');
/*!40000 ALTER TABLE `item_category` ENABLE KEYS */;


--
-- Definition of table `items`
--

DROP TABLE IF EXISTS `items`;
CREATE TABLE `items` (
  `item_id` int(11) NOT NULL AUTO_INCREMENT,
  `item_code` varchar(45) NOT NULL,
  `item_qty` double(10,2) NOT NULL,
  `item_price` double(10,2) NOT NULL,
  `dealers_price` double(10,2) DEFAULT NULL,
  `date_added` date DEFAULT NULL,
  `date_modified` date DEFAULT NULL,
  `manufacturers_id` int(10) unsigned DEFAULT NULL,
  `reorder_point` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`item_id`,`item_code`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=33 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `items`
--

/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` (`item_id`,`item_code`,`item_qty`,`item_price`,`dealers_price`,`date_added`,`date_modified`,`manufacturers_id`,`reorder_point`) VALUES 
 (17,'NLB',93.00,1580.00,1560.00,'2011-10-12','2011-11-13',4,0),
 (18,'NLP',94.00,985.00,960.00,'2011-10-12','2011-11-13',4,0),
 (19,'NSP',98.00,1275.00,1250.00,'2011-10-12','2011-11-13',0,0),
 (20,'NGP',100.00,1155.00,NULL,'2011-10-12','2011-10-23',0,0),
 (21,'NBP',98.50,1095.00,1075.00,'2011-10-12','2011-11-13',4,0),
 (22,'PSP',100.00,1085.00,NULL,'2011-10-12','2011-10-23',0,0),
 (23,'LSP',99.00,1185.00,NULL,'2011-10-12','2011-10-12',4,0),
 (25,'BFC',100.00,1195.00,NULL,'2011-10-12','2011-10-23',0,0),
 (26,'BSC',100.00,1260.00,NULL,'2011-10-12','2011-10-12',4,0),
 (27,'CBM 25',98.00,690.00,NULL,'2011-10-12','2011-10-23',0,0),
 (28,'CGM',200.00,1105.00,0.00,'2011-10-12','2011-11-13',0,0),
 (29,'item2',500.00,50.00,45.00,'2011-11-08','2011-11-13',0,0),
 (30,'item3',40.00,100.00,95.00,'2011-11-08','2011-11-08',0,0),
 (31,'item3',40.00,100.00,95.00,'2011-11-08','2011-11-08',0,0),
 (32,'item5',10.00,150.00,140.00,'2011-11-08','2011-11-08',0,0);
/*!40000 ALTER TABLE `items` ENABLE KEYS */;


--
-- Definition of table `items_description`
--

DROP TABLE IF EXISTS `items_description`;
CREATE TABLE `items_description` (
  `item_code` varchar(45) NOT NULL,
  `item_name` varchar(50) DEFAULT NULL,
  `item_description` varchar(100) DEFAULT NULL,
  `image` varchar(45) DEFAULT NULL,
  `status` tinyint(1) DEFAULT NULL,
  `unit_of_measure` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`item_code`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `items_description`
--

/*!40000 ALTER TABLE `items_description` DISABLE KEYS */;
INSERT INTO `items_description` (`item_code`,`item_name`,`item_description`,`image`,`status`,`unit_of_measure`) VALUES 
 ('BENG','BENG','BENG',NULL,1,'BAG'),
 ('BFC','Broiler Finisher Crumble','Broiler Finisher Crumble','',1,'BAG'),
 ('BSC','Broiler Starter Crumble','Broiler Starter Crumble','',1,'BAG'),
 ('CBM 25','Chick Booster Mash 25','Chick Booster Mash 25','',1,'BAG'),
 ('CGM','Chicken Grower Mash','Chicken Grower Mash','',1,'BAG'),
 ('CHRISTINE','CHRISTINE','CHRISTINE',NULL,1,'BAG'),
 ('CLM','Chicken Layer Mash','Chicken Layer Mash',NULL,1,'BAG'),
 ('CSM','Chick Starter Mash','Chich Starter Mash',NULL,1,'BAG'),
 ('FRYMASH','FRYMASH','FRYMASH',NULL,1,'BAG'),
 ('HFP','Hog Finisher Pellets','Hog Finisher Pellet',NULL,1,'BAG'),
 ('HGP','Hog Grower Pellet','Hog Grower Pellet',NULL,1,'BAG'),
 ('HSP','Hog Starter Crumble','Hog Starter Crumble',NULL,1,'BAG'),
 ('item2','item2','asdsadadsad','',1,'BAG'),
 ('item3','item3','asdasd','',1,'BAG'),
 ('item5','item5','adasddads','',1,'BAG'),
 ('LSP','Litter Saver Pellets','Litter Saver Pellets','',1,'BAG'),
 ('Milko Plus','Milko Plus','Milko Plus',NULL,1,'BAG'),
 ('NBP','Nutri Big Pellets','Nutri Big Pellets','',1,'BAG'),
 ('NGP','Nutri Gro Pellets','Nutri Gro Pellets','',1,'BAG'),
 ('NLB','Nutrilac Booster 1kg X 25','Nutrilac Booster 1kg X 25','',1,'BAG'),
 ('NLP','Nutrilac Pellets 1kg X 25','Nutrilac Pellets 1kg X 25','',1,'BAG'),
 ('NSP','Nutri Start Pellets','Nutri Start Pellets','',1,'BAG'),
 ('PBC','Premium Bangus Crumble','Premium Bangus Crumble',NULL,1,'BAG'),
 ('PBF','Premium Bangus Finisher ','Premium Bangus Finisher',NULL,1,'BAG'),
 ('PBG','Premium Bangus Grower ','Premium Bangus Grower',NULL,1,'BAG'),
 ('PBS','Premium Bangus Starter','Premium Bangus Starter',NULL,1,'BAG'),
 ('PDP (ORDINARY)','Pullet Developer Pellets (ORDINARY)','Pullet Developer Pellets (ORDINARY)',NULL,1,'BAG'),
 ('PDP (SPECIAL)','Pullet Developer Pellets ( SPECIAL)','Pullet Developer Pellets (SPECIAL)',NULL,1,'BAG'),
 ('PGB','Pigeon Pellet ','Pigeon Pellet',NULL,1,'BAG'),
 ('PGR','Pigeon Pellet  Red','Pigeon Pellet Red',NULL,1,'BAG'),
 ('PSP','Preg Sow Pelets','Preg Sow Pelets','',1,'BAG'),
 ('RED RICE','RED RICE','RED RICE',NULL,1,'BAG'),
 ('ROYAL CROWN','ROYAL CROWN','ROYAL CROWN',NULL,1,'BAG'),
 ('RPG','RPG','RPG',NULL,1,'BAG'),
 ('SBF','Surfer Bangus Finisher','Surfer Bangus Finisher',NULL,1,'BAG'),
 ('SBG','Surfer Bangus Grower','Sufer Bangus Grower',NULL,1,'BAG'),
 ('SBS','Surfer Bangus Starter','Surfer Bangus Starter',NULL,1,'BAG'),
 ('SPS PRE-STARTER','Sps Pre-Starter','Sps Pre-starter',NULL,1,'BAG'),
 ('TAHOP','TAHOP','TAHOP',NULL,1,'BAG'),
 ('THB BOOSTER','THB BOOSTER','THB BOOSTER',NULL,1,'CASE'),
 ('THB DEVELOPER ','THB DEVELOPER','THB DEVELOPER',NULL,1,'CASE'),
 ('THB ENERTONE','THB ENERTONE','THB ENERTONE',NULL,1,'CASE'),
 ('THB FINISHER','THB FINISHER ','THB FINISHER',NULL,1,'CASE'),
 ('THB HI-LANDER','THB HI-LANDER','THB HI-LANDER',NULL,1,'CASE'),
 ('THB MULTIGRAIN','THB MULTIGRAIN','THB MULTIGRAIN',NULL,1,'CASE'),
 ('THB PLATINUM ','THB PLATINUM','THB PLATINUN',NULL,1,'CASE'),
 ('THB SUCCESSOR','THB SUCCESSOR','THB SUCCESSOR',NULL,1,'CASE'),
 ('TIKI-TIKI','TIKI-TIKI','TIKI-TIKI',NULL,1,'BAG');
/*!40000 ALTER TABLE `items_description` ENABLE KEYS */;


--
-- Definition of table `last_inventory`
--

DROP TABLE IF EXISTS `last_inventory`;
CREATE TABLE `last_inventory` (
  `item_id` int(10) unsigned DEFAULT NULL,
  `item_code` varchar(45) DEFAULT NULL,
  `beginning_balance` double(10,2) DEFAULT NULL,
  `ending_balance` double(10,2) DEFAULT NULL,
  `date` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `last_inventory`
--

/*!40000 ALTER TABLE `last_inventory` DISABLE KEYS */;
/*!40000 ALTER TABLE `last_inventory` ENABLE KEYS */;


--
-- Definition of table `manufacturers`
--

DROP TABLE IF EXISTS `manufacturers`;
CREATE TABLE `manufacturers` (
  `manufacturers_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `manufacturers_name` varchar(45) NOT NULL,
  `manufacturers_add` varchar(45) NOT NULL,
  `manufacturers_number` varchar(45) NOT NULL,
  PRIMARY KEY (`manufacturers_id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `manufacturers`
--

/*!40000 ALTER TABLE `manufacturers` DISABLE KEYS */;
INSERT INTO `manufacturers` (`manufacturers_id`,`manufacturers_name`,`manufacturers_add`,`manufacturers_number`) VALUES 
 (4,'UNIFEEDS','CEBU',''),
 (5,'UNIVET','Mandaluyong City, Philippines','');
/*!40000 ALTER TABLE `manufacturers` ENABLE KEYS */;


--
-- Definition of table `municipal_agent`
--

DROP TABLE IF EXISTS `municipal_agent`;
CREATE TABLE `municipal_agent` (
  `agent_id` int(10) unsigned DEFAULT NULL,
  `municipal_id` varchar(45) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `municipal_agent`
--

/*!40000 ALTER TABLE `municipal_agent` DISABLE KEYS */;
INSERT INTO `municipal_agent` (`agent_id`,`municipal_id`) VALUES 
 (1,'1'),
 (1,'2'),
 (1,'3'),
 (1,'4'),
 (2,'5'),
 (2,'7'),
 (2,'8'),
 (3,'9'),
 (3,'10'),
 (3,'11'),
 (3,'12'),
 (4,'13'),
 (5,'6');
/*!40000 ALTER TABLE `municipal_agent` ENABLE KEYS */;


--
-- Definition of table `municipalities`
--

DROP TABLE IF EXISTS `municipalities`;
CREATE TABLE `municipalities` (
  `municipal_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `municipal_name` varchar(45) NOT NULL,
  PRIMARY KEY (`municipal_id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=14 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `municipalities`
--

/*!40000 ALTER TABLE `municipalities` DISABLE KEYS */;
INSERT INTO `municipalities` (`municipal_id`,`municipal_name`) VALUES 
 (1,'CALAPE'),
 (2,'TUBIGON'),
 (3,'CLARIN'),
 (4,'LOON'),
 (5,'UBAY'),
 (6,'SAGBAYAN'),
 (7,'BIEN UNIDO'),
 (8,'TRINIDAD'),
 (9,'TALIBON'),
 (10,'INABANGGA'),
 (11,'BUENAVISTA'),
 (12,'JETAFE'),
 (13,'CARMEN');
/*!40000 ALTER TABLE `municipalities` ENABLE KEYS */;


--
-- Definition of table `payment`
--

DROP TABLE IF EXISTS `payment`;
CREATE TABLE `payment` (
  `payment_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `customer_id` int(10) unsigned DEFAULT NULL,
  `amount` double(2,2) NOT NULL,
  `date_of_payment` datetime NOT NULL,
  `remarks` varchar(45) NOT NULL,
  PRIMARY KEY (`payment_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `payment`
--

/*!40000 ALTER TABLE `payment` DISABLE KEYS */;
/*!40000 ALTER TABLE `payment` ENABLE KEYS */;


--
-- Definition of table `payment_records`
--

DROP TABLE IF EXISTS `payment_records`;
CREATE TABLE `payment_records` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `sales_order_no` varchar(45) DEFAULT NULL,
  `amount` double(10,2) DEFAULT NULL,
  `balance` double(10,2) DEFAULT NULL,
  `payment_date` datetime DEFAULT NULL,
  `remarks` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `payment_records`
--

/*!40000 ALTER TABLE `payment_records` DISABLE KEYS */;
INSERT INTO `payment_records` (`id`,`sales_order_no`,`amount`,`balance`,`payment_date`,`remarks`) VALUES 
 (1,'SO-0000052',1000.00,3550.00,'2011-11-06 00:00:00',''),
 (2,'SO-0000052',3000.00,1550.00,'2011-11-06 00:00:00',''),
 (3,'SO-0000055',1000.00,1565.00,'2011-11-13 00:00:00',''),
 (4,'SO-0000057',0.00,2510.00,'2011-11-13 00:00:00',''),
 (5,'SO-0000057',1000.00,1510.00,'2011-11-13 00:00:00',''),
 (6,'SO-0000057',1000.00,1510.00,'2011-11-13 00:00:00',''),
 (7,'SO-0000057',510.00,2000.00,'2011-11-13 00:00:00','fully paid');
/*!40000 ALTER TABLE `payment_records` ENABLE KEYS */;


--
-- Definition of table `salesorder_responsible`
--

DROP TABLE IF EXISTS `salesorder_responsible`;
CREATE TABLE `salesorder_responsible` (
  `SaleOrder_Responsible_Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Prepared_by` varchar(100) DEFAULT NULL,
  `Checked_by` varchar(100) DEFAULT NULL,
  `Posted_by` varchar(100) DEFAULT NULL,
  `Delivered_by` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`SaleOrder_Responsible_Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `salesorder_responsible`
--

/*!40000 ALTER TABLE `salesorder_responsible` DISABLE KEYS */;
INSERT INTO `salesorder_responsible` (`SaleOrder_Responsible_Id`,`Prepared_by`,`Checked_by`,`Posted_by`,`Delivered_by`) VALUES 
 (1,'jhun','hammerj','kissha','milky');
/*!40000 ALTER TABLE `salesorder_responsible` ENABLE KEYS */;


--
-- Definition of table `stock_in`
--

DROP TABLE IF EXISTS `stock_in`;
CREATE TABLE `stock_in` (
  `stockin_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `item_id` varchar(45) NOT NULL,
  `qty_in` int(10) unsigned NOT NULL,
  PRIMARY KEY (`stockin_id`)
) ENGINE=InnoDB AUTO_INCREMENT=42 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_in`
--

/*!40000 ALTER TABLE `stock_in` DISABLE KEYS */;
INSERT INTO `stock_in` (`stockin_id`,`item_id`,`qty_in`) VALUES 
 (3,'8',10),
 (4,'8',100),
 (5,'8',100),
 (6,'10',20),
 (7,'8',100),
 (8,'8',100),
 (9,'8',100),
 (10,'8',500),
 (11,'10',500),
 (12,'8',8),
 (13,'8',5),
 (14,'1',9),
 (15,'8',78),
 (16,'8',10),
 (17,'8',10),
 (18,'11',50),
 (19,'11',50),
 (20,'12',2),
 (21,'14',2),
 (22,'12',5),
 (23,'15',5),
 (24,'19',200),
 (25,'21',200),
 (26,'20',100),
 (27,'22',200),
 (28,'23',1500),
 (29,'27',200),
 (30,'17',100),
 (31,'23',100),
 (32,'22',100),
 (33,'27',100),
 (34,'28',100),
 (35,'21',150),
 (36,'20',50),
 (37,'22',50),
 (38,'23',50),
 (39,'27',20),
 (40,'17',10),
 (41,'17',500);
/*!40000 ALTER TABLE `stock_in` ENABLE KEYS */;


--
-- Definition of table `stock_in_reference`
--

DROP TABLE IF EXISTS `stock_in_reference`;
CREATE TABLE `stock_in_reference` (
  `reference_no` varchar(45) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_in_reference`
--

/*!40000 ALTER TABLE `stock_in_reference` DISABLE KEYS */;
INSERT INTO `stock_in_reference` (`reference_no`) VALUES 
 ('17');
/*!40000 ALTER TABLE `stock_in_reference` ENABLE KEYS */;


--
-- Definition of table `stock_in_transaction`
--

DROP TABLE IF EXISTS `stock_in_transaction`;
CREATE TABLE `stock_in_transaction` (
  `stock_in_transaction_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `reference_no` varchar(45) NOT NULL,
  `stocked_in_to` varchar(45) DEFAULT NULL,
  `from_supplier` int(10) unsigned DEFAULT NULL,
  `remarks` text,
  `stock_in_date` date DEFAULT NULL,
  `total_number_of_items` int(10) unsigned DEFAULT NULL,
  `total_qty` int(10) unsigned DEFAULT NULL,
  `prepared_by` varchar(45) DEFAULT NULL,
  `approved_by` varchar(45) DEFAULT NULL,
  `received_by` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`stock_in_transaction_id`)
) ENGINE=InnoDB AUTO_INCREMENT=28 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_in_transaction`
--

/*!40000 ALTER TABLE `stock_in_transaction` DISABLE KEYS */;
INSERT INTO `stock_in_transaction` (`stock_in_transaction_id`,`reference_no`,`stocked_in_to`,`from_supplier`,`remarks`,`stock_in_date`,`total_number_of_items`,`total_qty`,`prepared_by`,`approved_by`,`received_by`) VALUES 
 (5,'SI-000001','WH-02 STOCKROOM(BODEGA)',1,'asdsadad','2011-10-09',1,10,'obeng','',''),
 (6,'SI-000001','WH-02 STOCKROOM(BODEGA)',1,'','2011-10-09',1,100,'','',''),
 (7,'SI-000001','WH-02 STOCKROOM(BODEGA)',1,'','2011-10-09',1,100,'','',''),
 (8,'SI-000001','WH-02 STOCKROOM(BODEGA)',1,'','2011-10-09',1,20,'','',''),
 (9,'SI-000001','WH-02 STOCKROOM(BODEGA)',2,'sample','2011-10-09',1,100,'obeng','',''),
 (10,'SI-000001','WH-02 STOCKROOM(BODEGA)',2,'samnple','2011-10-09',1,100,'obeng','',''),
 (11,'SI-000001','WH-02 STOCKROOM(BODEGA)',2,'sample','2011-10-09',1,100,'obeng','',''),
 (12,'SI-000001','WH-02 STOCKROOM(BODEGA)',1,'','2011-10-09',2,1000,'','',''),
 (13,'SI-000002','WH-02 STOCKROOM(BODEGA)',2,'dssd','2011-10-09',1,8,'sdfs','sdfs','sdfs'),
 (14,'SI-000003','WH-02 STOCKROOM(BODEGA)',2,'sdfasdf','2011-10-09',2,14,'asfsdfsa','sdfsda','sdfs'),
 (15,'SI-000004','WH-02 STOCKROOM(BODEGA)',2,'dgd','2011-10-09',1,78,'','',''),
 (16,'SI-000005','WH-02 STOCKROOM(BODEGA)',2,'aasdsdsadd','2011-10-09',1,10,'asdsad','asd',''),
 (17,'SI-000006','WH-02 STOCKROOM(BODEGA)',2,'asdasd','2011-10-09',1,10,'asdasd','',''),
 (18,'SI-000007','WH-02 STOCKROOM(BODEGA)',3,'vbmkgi','2011-10-09',1,50,'','',''),
 (19,'SI-000008','WH-02 STOCKROOM(BODEGA)',3,'','2011-10-09',1,50,'','',''),
 (20,'SI-000009','WH-02 STOCKROOM(BODEGA)',4,'FROM: CEBU / WING VAN  / BOY','2011-10-10',2,4,'BENG','SIR AHOC','RAMBO'),
 (21,'SI-0000010','BODEGA / TINDAHAN',4,'PICK UP CEBU / WING VAN - BOY ','2011-10-10',2,10,'BENG','SIR AHOC','RAMBO'),
 (22,'SI-0000011','WH-02 STOCKROOM(BODEGA)',4,'PICK UP / WING VAN - BOY','2011-10-17',5,2200,'                              BENG','                              SIR AHOC','                                RAMBO'),
 (23,'SI-0000012','WH-02 STOCKROOM(BODEGA)',4,'','2011-10-17',1,200,'','',''),
 (24,'SI-0000013','WH-02 STOCKROOM(BODEGA)',4,'pick up/ wing van - boy','2011-10-23',5,500,'     beng','     sir ahoc','          ramboo'),
 (25,'SI-0000014','WH-01 STOCKROOM(KATUNGAN)',4,'SHIPPED / WING VAN - ANTONIO','2011-10-28',5,320,'                BENG','                 RAMBO','           SIR AHOC'),
 (26,'SI-0000015','WH-02 STOCKROOM(BODEGA)',4,'','2011-11-01',1,10,'obeng','hehe','ahoc'),
 (27,'SI-0000016','WH-02 STOCKROOM(BODEGA)',4,'ok ra ni','2011-11-05',1,500,'ham','jhun','cool');
/*!40000 ALTER TABLE `stock_in_transaction` ENABLE KEYS */;


--
-- Definition of table `stock_in_transaction_to_stock_in_items`
--

DROP TABLE IF EXISTS `stock_in_transaction_to_stock_in_items`;
CREATE TABLE `stock_in_transaction_to_stock_in_items` (
  `stock_in_transaction_id` int(10) unsigned DEFAULT NULL,
  `stock_id` int(10) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_in_transaction_to_stock_in_items`
--

/*!40000 ALTER TABLE `stock_in_transaction_to_stock_in_items` DISABLE KEYS */;
INSERT INTO `stock_in_transaction_to_stock_in_items` (`stock_in_transaction_id`,`stock_id`) VALUES 
 (7,0),
 (8,0),
 (10,0),
 (11,9),
 (12,10),
 (12,11),
 (13,12),
 (14,13),
 (14,14),
 (15,15),
 (16,16),
 (17,17),
 (18,18),
 (19,19),
 (20,20),
 (20,21),
 (21,22),
 (21,23),
 (22,24),
 (22,25),
 (22,26),
 (22,27),
 (22,28),
 (23,29),
 (24,30),
 (24,31),
 (24,32),
 (24,33),
 (24,34),
 (25,35),
 (25,36),
 (25,37),
 (25,38),
 (25,39),
 (26,40),
 (27,41);
/*!40000 ALTER TABLE `stock_in_transaction_to_stock_in_items` ENABLE KEYS */;


--
-- Definition of table `stock_out`
--

DROP TABLE IF EXISTS `stock_out`;
CREATE TABLE `stock_out` (
  `stockout_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `item_id` varchar(45) DEFAULT NULL,
  `qty_out` double(10,2) DEFAULT NULL,
  `amount` double(10,2) DEFAULT NULL,
  `discount` double(10,2) DEFAULT NULL,
  `tracking_price` double(10,2) DEFAULT NULL,
  PRIMARY KEY (`stockout_id`)
) ENGINE=InnoDB AUTO_INCREMENT=22 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_out`
--

/*!40000 ALTER TABLE `stock_out` DISABLE KEYS */;
INSERT INTO `stock_out` (`stockout_id`,`item_id`,`qty_out`,`amount`,`discount`,`tracking_price`) VALUES 
 (1,'17',1.00,1580.00,NULL,NULL),
 (2,'17',1.00,1580.00,NULL,NULL),
 (3,'18',1.00,985.00,NULL,NULL),
 (4,'19',1.00,1275.00,NULL,NULL),
 (5,'17',1.00,1580.00,NULL,NULL),
 (6,'18',1.00,985.00,NULL,NULL),
 (7,'19',1.00,1275.00,NULL,NULL),
 (8,'17',1.00,1580.00,NULL,NULL),
 (9,'27',1.00,690.00,NULL,NULL),
 (10,'17',1.00,1580.00,NULL,NULL),
 (11,'27',1.00,690.00,NULL,NULL),
 (12,'21',1.00,1095.00,NULL,NULL),
 (13,'23',1.00,1185.00,NULL,NULL),
 (14,'18',1.00,985.00,0.00,0.00),
 (15,'21',0.50,537.50,0.00,0.00),
 (16,'17',1.00,1560.00,10.00,10.00),
 (17,'18',1.00,965.00,5.00,10.00),
 (18,'17',1.00,1555.00,10.00,5.00),
 (19,'18',1.00,955.00,10.00,5.00),
 (20,'17',1.00,1555.00,10.00,5.00),
 (21,'18',1.00,955.00,10.00,5.00);
/*!40000 ALTER TABLE `stock_out` ENABLE KEYS */;


--
-- Definition of table `stock_out_reference`
--

DROP TABLE IF EXISTS `stock_out_reference`;
CREATE TABLE `stock_out_reference` (
  `reference_no` int(10) unsigned NOT NULL DEFAULT '0'
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_out_reference`
--

/*!40000 ALTER TABLE `stock_out_reference` DISABLE KEYS */;
INSERT INTO `stock_out_reference` (`reference_no`) VALUES 
 (58);
/*!40000 ALTER TABLE `stock_out_reference` ENABLE KEYS */;


--
-- Definition of table `stock_out_transaction`
--

DROP TABLE IF EXISTS `stock_out_transaction`;
CREATE TABLE `stock_out_transaction` (
  `sales_order_no` varchar(45) NOT NULL,
  `responsible_customer` int(10) unsigned DEFAULT NULL,
  `responsible_agent` int(10) unsigned DEFAULT NULL,
  `discount` double(10,2) NOT NULL,
  `grand_total` double(10,2) NOT NULL,
  `net_total` double(10,2) NOT NULL,
  `tendered_amount` double(10,2) DEFAULT NULL,
  `change` double(10,2) DEFAULT NULL,
  `delivery_date` datetime NOT NULL,
  PRIMARY KEY (`sales_order_no`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_out_transaction`
--

/*!40000 ALTER TABLE `stock_out_transaction` DISABLE KEYS */;
INSERT INTO `stock_out_transaction` (`sales_order_no`,`responsible_customer`,`responsible_agent`,`discount`,`grand_total`,`net_total`,`tendered_amount`,`change`,`delivery_date`) VALUES 
 ('SO-000001',9,3,0.00,1580.00,1580.00,NULL,NULL,'2011-11-01 20:34:10'),
 ('SO-0000010',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 10:54:48'),
 ('SO-0000011',NULL,NULL,0.00,1580.00,1580.00,0.00,0.00,'2011-11-05 11:01:38'),
 ('SO-0000012',9,3,0.00,1580.00,1580.00,0.00,0.00,'2011-11-05 11:37:40'),
 ('SO-0000013',9,3,0.00,1580.00,1580.00,0.00,0.00,'2011-11-05 11:38:23'),
 ('SO-0000014',NULL,NULL,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:40:08'),
 ('SO-0000015',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:41:40'),
 ('SO-0000016',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:43:50'),
 ('SO-0000017',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:44:24'),
 ('SO-0000018',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:45:35'),
 ('SO-0000019',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:47:42'),
 ('SO-000002',9,3,0.00,1580.00,1580.00,0.00,0.00,'2011-11-04 00:28:21'),
 ('SO-0000020',9,3,0.00,985.00,985.00,0.00,0.00,'2011-11-05 11:56:10'),
 ('SO-0000021',9,3,0.00,1275.00,1275.00,0.00,0.00,'2011-11-05 12:16:22'),
 ('SO-0000022',9,3,0.00,3960.00,3960.00,0.00,0.00,'2011-11-05 12:17:43'),
 ('SO-0000023',9,3,0.00,2360.00,2360.00,0.00,0.00,'2011-11-05 12:18:47'),
 ('SO-0000024',9,3,0.00,1580.00,1580.00,2000.00,0.00,'2011-11-05 13:17:46'),
 ('SO-0000025',NULL,NULL,0.00,1580.00,1580.00,2000.00,0.00,'2011-11-05 13:19:08'),
 ('SO-0000026',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 13:19:42'),
 ('SO-0000027',NULL,NULL,0.00,1580.00,1580.00,222222.00,220642.00,'2011-11-05 20:19:12'),
 ('SO-0000028',14,NULL,0.00,1580.00,1580.00,10000.00,8420.00,'2011-11-05 20:35:20'),
 ('SO-0000029',NULL,NULL,0.00,1580.00,1580.00,200000.00,198420.00,'2011-11-05 20:43:18'),
 ('SO-000003',9,3,0.00,1580.00,1580.00,0.00,0.00,'2011-11-04 00:31:06'),
 ('SO-0000030',0,0,0.00,1580.00,1580.00,199999.00,198419.00,'2011-11-05 21:10:11'),
 ('SO-0000031',0,0,0.00,1580.00,1580.00,99999.00,98419.00,'2011-11-05 21:11:16'),
 ('SO-0000032',0,0,0.00,1580.00,1580.00,88888.00,87308.00,'2011-11-05 21:12:23'),
 ('SO-0000033',0,0,0.00,1580.00,1580.00,6666.00,5086.00,'2011-11-05 21:13:39'),
 ('SO-0000034',0,0,0.00,1580.00,1580.00,9999.00,8419.00,'2011-11-05 21:19:39'),
 ('SO-0000035',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 21:58:19'),
 ('SO-0000036',18,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:13:10'),
 ('SO-0000037',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:27:20'),
 ('SO-0000038',9,3,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:31:54'),
 ('SO-0000039',9,3,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:34:12'),
 ('SO-0000040',17,NULL,0.00,1580.00,1580.00,1580.00,0.00,'2011-11-05 23:39:29'),
 ('SO-0000041',11,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:40:27'),
 ('SO-0000042',NULL,NULL,0.00,1580.00,1580.00,0.00,-1580.00,'2011-11-05 23:45:14'),
 ('SO-0000043',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:45:41'),
 ('SO-0000044',13,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:47:36'),
 ('SO-0000045',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:51:07'),
 ('SO-0000046',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-05 23:54:41'),
 ('SO-0000047',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-06 00:24:27'),
 ('SO-0000048',NULL,NULL,0.00,1580.00,1580.00,2000.00,420.00,'2011-11-06 00:26:53'),
 ('SO-0000049',9,3,0.00,3840.00,3840.00,3840.00,0.00,'2011-11-06 00:28:13'),
 ('SO-0000050',9,3,0.00,3840.00,3840.00,4000.00,160.00,'2011-11-06 00:29:35'),
 ('SO-0000051',13,NULL,0.00,2270.00,2270.00,2270.00,0.00,'2011-11-06 00:37:37'),
 ('SO-0000052',8,NULL,0.00,4550.00,4550.00,0.00,-4550.00,'2011-11-06 02:19:26'),
 ('SO-0000053',9,1,0.00,985.00,985.00,1000.00,15.00,'2011-11-12 18:16:50'),
 ('SO-0000054',8,3,0.00,537.50,537.50,600.00,62.50,'2011-11-13 14:50:58'),
 ('SO-0000055',8,3,0.00,2525.00,2525.00,0.00,-2525.00,'2011-11-13 15:30:42'),
 ('SO-0000056',8,3,0.00,2510.00,2510.00,0.00,-2510.00,'2011-11-13 15:51:28'),
 ('SO-0000057',8,3,0.00,2510.00,2510.00,0.00,-2510.00,'2011-11-13 15:56:16'),
 ('SO-000006',NULL,NULL,0.00,985.00,985.00,0.00,0.00,'2011-11-04 00:36:03'),
 ('SO-000008',9,3,0.00,1580.00,1580.00,0.00,0.00,'2011-11-04 00:40:35'),
 ('SO-000009',NULL,NULL,0.00,1580.00,1580.00,0.00,0.00,'2011-11-04 00:46:30');
/*!40000 ALTER TABLE `stock_out_transaction` ENABLE KEYS */;


--
-- Definition of table `stock_out_transaction_stock_out_items`
--

DROP TABLE IF EXISTS `stock_out_transaction_stock_out_items`;
CREATE TABLE `stock_out_transaction_stock_out_items` (
  `sales_order_no` varchar(45) DEFAULT NULL,
  `stockout_id` int(10) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_out_transaction_stock_out_items`
--

/*!40000 ALTER TABLE `stock_out_transaction_stock_out_items` DISABLE KEYS */;
INSERT INTO `stock_out_transaction_stock_out_items` (`sales_order_no`,`stockout_id`) VALUES 
 ('SO-0000051',8),
 ('SO-0000051',9),
 ('SO-0000052',10),
 ('SO-0000052',11),
 ('SO-0000052',12),
 ('SO-0000052',13),
 ('SO-0000053',14),
 ('SO-0000054',15),
 ('SO-0000055',16),
 ('SO-0000055',17),
 ('SO-0000056',18),
 ('SO-0000056',19),
 ('SO-0000057',20),
 ('SO-0000057',21);
/*!40000 ALTER TABLE `stock_out_transaction_stock_out_items` ENABLE KEYS */;


--
-- Definition of table `temp`
--

DROP TABLE IF EXISTS `temp`;
CREATE TABLE `temp` (
  `item_id` int(10) unsigned DEFAULT NULL,
  `item_code` varchar(45) DEFAULT NULL,
  `ending_balance` double(10,2) DEFAULT NULL,
  `item_qty` double(10,2) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `temp`
--

/*!40000 ALTER TABLE `temp` DISABLE KEYS */;
/*!40000 ALTER TABLE `temp` ENABLE KEYS */;


--
-- Definition of table `useraccount`
--

DROP TABLE IF EXISTS `useraccount`;
CREATE TABLE `useraccount` (
  `username` varchar(50) NOT NULL,
  `password` varchar(50) NOT NULL,
  `user_type` varchar(50) NOT NULL DEFAULT 'user',
  PRIMARY KEY (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `useraccount`
--

/*!40000 ALTER TABLE `useraccount` DISABLE KEYS */;
INSERT INTO `useraccount` (`username`,`password`,`user_type`) VALUES 
 ('admin','21232f297a57a5a743894a0e4a801fc3','admin'),
 ('user','4297f44b13955235245b2497399d7a93','user');
/*!40000 ALTER TABLE `useraccount` ENABLE KEYS */;


--
-- Definition of view `view_ending_balance`
--

DROP TABLE IF EXISTS `view_ending_balance`;
DROP VIEW IF EXISTS `view_ending_balance`;
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `view_ending_balance` AS select `items`.`item_code` AS `item_code`,`items`.`item_qty` AS `item_qty` from `items`;



/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;

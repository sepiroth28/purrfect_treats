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
-- Definition of table `account_recievable`
--

DROP TABLE IF EXISTS `account_recievable`;
CREATE TABLE `account_recievable` (
  `account_recievable_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `gross_amount` double(2,2) NOT NULL,
  `total_discount_amount` double(2,2) NOT NULL,
  `net_total` double(2,2) NOT NULL,
  `acount_recievable_date` datetime NOT NULL,
  `user_id` int(10) unsigned NOT NULL,
  `customer_id` int(10) unsigned NOT NULL,
  `status` varchar(45) NOT NULL,
  PRIMARY KEY (`account_recievable_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `account_recievable`
--

/*!40000 ALTER TABLE `account_recievable` DISABLE KEYS */;
/*!40000 ALTER TABLE `account_recievable` ENABLE KEYS */;


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
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `agent`
--

/*!40000 ALTER TABLE `agent` DISABLE KEYS */;
INSERT INTO `agent` (`agent_id`,`Name`,`Mobile`,`address`) VALUES 
 (1,'aaa','aa','adress'),
 (2,'aaa','aa','ad'),
 (3,'assdf','sdfsfsdf','adr'),
 (4,'tokoy','0910-2525-2525','');
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
-- Definition of table `customers`
--

DROP TABLE IF EXISTS `customers`;
CREATE TABLE `customers` (
  `customers_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `customers_name` varchar(45) NOT NULL,
  `customers_add` varchar(45) NOT NULL,
  `customers_number` varchar(45) NOT NULL,
  `municipal_id` int(10) unsigned NOT NULL,
  PRIMARY KEY (`customers_id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `customers`
--

/*!40000 ALTER TABLE `customers` DISABLE KEYS */;
INSERT INTO `customers` (`customers_id`,`customers_name`,`customers_add`,`customers_number`,`municipal_id`) VALUES 
 (1,'aris','tagbdd','1234',0),
 (2,'Hammerj Merto ','Cebu','123456789',0),
 (4,'jhun','jhun','1',0),
 (5,'ham','ham','123456',0),
 (7,'jun','capitol hills','123132sdffds',0);
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
  `discount_code` varchar(45) NOT NULL,
  `discount_name` varchar(45) NOT NULL,
  `percentage` varchar(45) NOT NULL,
  `percentage_amount` double(2,2) NOT NULL,
  PRIMARY KEY (`discount_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `discount`
--

/*!40000 ALTER TABLE `discount` DISABLE KEYS */;
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
) ENGINE=InnoDB AUTO_INCREMENT=33 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `inventory`
--

/*!40000 ALTER TABLE `inventory` DISABLE KEYS */;
INSERT INTO `inventory` (`id`,`item_id`,`item_code`,`beginning_balance`,`ending_balance`,`date`) VALUES 
 (1,1,'001',800.00,800.00,'2011-10-22 00:00:00'),
 (2,2,'002',3540.00,3540.00,'2011-10-22 00:00:00'),
 (3,8,'ham001',3452.00,3452.00,'2011-10-22 00:00:00'),
 (4,11,'milk001',617.00,617.00,'2011-10-22 00:00:00'),
 (8,1,'001',800.00,800.00,'2011-10-22 13:28:09'),
 (9,2,'002',3540.00,3540.00,'2011-10-22 13:28:09'),
 (10,8,'ham001',3452.00,3452.00,'2011-10-22 13:28:09'),
 (11,11,'milk001',617.00,617.00,'2011-10-22 13:28:09'),
 (15,1,'001',800.00,800.00,'2011-10-22 13:28:41'),
 (16,2,'002',3540.00,3540.00,'2011-10-22 13:28:41'),
 (17,8,'ham001',3452.00,3452.00,'2011-10-22 13:28:41'),
 (18,11,'milk001',617.00,617.00,'2011-10-22 13:28:41'),
 (22,1,'001',800.00,800.00,'2011-10-22 13:29:07'),
 (23,2,'002',3540.00,3540.00,'2011-10-22 13:29:07'),
 (24,8,'ham001',3452.00,3452.00,'2011-10-22 13:29:07'),
 (25,11,'milk001',617.00,617.00,'2011-10-22 13:29:07'),
 (29,1,'001',800.00,800.00,'2011-10-22 13:32:17'),
 (30,2,'002',3540.00,3540.00,'2011-10-22 13:32:17'),
 (31,8,'ham001',3452.00,3452.00,'2011-10-22 13:32:17'),
 (32,11,'milk001',617.00,617.00,'2011-10-22 13:32:17');
/*!40000 ALTER TABLE `inventory` ENABLE KEYS */;


--
-- Definition of table `items`
--

DROP TABLE IF EXISTS `items`;
CREATE TABLE `items` (
  `item_id` int(11) NOT NULL AUTO_INCREMENT,
  `item_code` varchar(45) NOT NULL,
  `item_qty` double(10,2) NOT NULL,
  `item_price` double(10,2) NOT NULL,
  `date_added` date DEFAULT NULL,
  `date_modified` date DEFAULT NULL,
  `manufacturers_id` int(10) unsigned DEFAULT NULL,
  `reorder_point` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`item_id`,`item_code`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `items`
--

/*!40000 ALTER TABLE `items` DISABLE KEYS */;
INSERT INTO `items` (`item_id`,`item_code`,`item_qty`,`item_price`,`date_added`,`date_modified`,`manufacturers_id`,`reorder_point`) VALUES 
 (1,'001',800.00,822.00,'2011-09-16','2011-10-09',0,10),
 (2,'002',3540.00,20.00,'2011-09-16','2011-09-16',3,10),
 (8,'ham001',3452.00,1500.00,'2011-09-29','2011-10-06',0,15),
 (11,'milk001',617.00,550.00,'2011-10-09','2011-10-09',0,80);
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
 ('001','chicken feeds','aaaa',NULL,1,'sack'),
 ('002','asdjsakd','asdasd','',1,'kilo'),
 ('ham001','ham001','ham for ham','',1,'kilo'),
 ('milk001','bear brand','milk bear brands','',1,'sack');
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
INSERT INTO `last_inventory` (`item_id`,`item_code`,`beginning_balance`,`ending_balance`,`date`) VALUES 
 (1,'001',800.00,800.00,'2011-10-22 13:32:17'),
 (2,'002',3540.00,3540.00,'2011-10-22 13:32:17'),
 (8,'ham001',3452.00,3452.00,'2011-10-22 13:32:17'),
 (11,'milk001',617.00,617.00,'2011-10-22 13:32:17');
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
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `manufacturers`
--

/*!40000 ALTER TABLE `manufacturers` DISABLE KEYS */;
INSERT INTO `manufacturers` (`manufacturers_id`,`manufacturers_name`,`manufacturers_add`,`manufacturers_number`) VALUES 
 (3,'BMEG','cebu','123456');
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
 (4,'1');
/*!40000 ALTER TABLE `municipal_agent` ENABLE KEYS */;


--
-- Definition of table `municipalities`
--

DROP TABLE IF EXISTS `municipalities`;
CREATE TABLE `municipalities` (
  `municipal_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `municipal_name` varchar(45) NOT NULL,
  PRIMARY KEY (`municipal_id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `municipalities`
--

/*!40000 ALTER TABLE `municipalities` DISABLE KEYS */;
INSERT INTO `municipalities` (`municipal_id`,`municipal_name`) VALUES 
 (1,'Albur'),
 (2,'Baclayon'),
 (3,'Tagbilaran');
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
-- Definition of table `sales`
--

DROP TABLE IF EXISTS `sales`;
CREATE TABLE `sales` (
  `sales_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `gross_amount` double(2,2) NOT NULL,
  `total_discount_amount` double(2,2) NOT NULL,
  `net_total` double(2,2) NOT NULL,
  `sales_date` datetime NOT NULL,
  `user_id` int(10) unsigned NOT NULL,
  `sold_to` varchar(45) NOT NULL,
  PRIMARY KEY (`sales_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `sales`
--

/*!40000 ALTER TABLE `sales` DISABLE KEYS */;
/*!40000 ALTER TABLE `sales` ENABLE KEYS */;


--
-- Definition of table `sales_cart`
--

DROP TABLE IF EXISTS `sales_cart`;
CREATE TABLE `sales_cart` (
  `sales_id` int(10) unsigned NOT NULL DEFAULT '0',
  `cart_id` int(10) unsigned DEFAULT NULL,
  PRIMARY KEY (`sales_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `sales_cart`
--

/*!40000 ALTER TABLE `sales_cart` DISABLE KEYS */;
/*!40000 ALTER TABLE `sales_cart` ENABLE KEYS */;


--
-- Definition of table `stock_in`
--

DROP TABLE IF EXISTS `stock_in`;
CREATE TABLE `stock_in` (
  `stockin_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `item_id` varchar(45) NOT NULL,
  `qty_in` int(10) unsigned NOT NULL,
  PRIMARY KEY (`stockin_id`)
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=latin1;

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
 (18,'11',67);
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
 ('8');
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
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=latin1;

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
 (18,'SI-000007','WH-02 STOCKROOM(BODEGA)',3,'cvxcv','2011-10-09',1,67,'cxxcv','cxvxc','xcvxc');
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
 (18,18);
/*!40000 ALTER TABLE `stock_in_transaction_to_stock_in_items` ENABLE KEYS */;


--
-- Definition of table `stock_out`
--

DROP TABLE IF EXISTS `stock_out`;
CREATE TABLE `stock_out` (
  `stockout_id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `transaction_type` varchar(45) NOT NULL,
  `affected_id` varchar(45) NOT NULL,
  `stockout_date` datetime NOT NULL,
  PRIMARY KEY (`stockout_id`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `stock_out`
--

/*!40000 ALTER TABLE `stock_out` DISABLE KEYS */;
/*!40000 ALTER TABLE `stock_out` ENABLE KEYS */;


--
-- Definition of table `temp`
--

DROP TABLE IF EXISTS `temp`;
CREATE TABLE `temp` (
  `item_id` int(10) unsigned DEFAULT NULL,
  `item_code` varchar(45) DEFAULT NULL,
  `ending_balance` double(10,2) DEFAULT NULL,
  `item_qty` double(10,2)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Dumping data for table `temp`
--

/*!40000 ALTER TABLE `temp` DISABLE KEYS */;
INSERT INTO `temp` (`item_id`,`item_code`,`ending_balance`,`item_qty`) VALUES 
 (1,'001',800.00,800.00),
 (2,'002',3540.00,3540.00),
 (8,'ham001',3452.00,3452.00),
 (11,'milk001',617.00,617.00);
/*!40000 ALTER TABLE `temp` ENABLE KEYS */;




/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;

Attribute VB_Name = "Entities"
Public Const rebates_table As String = "DROP TABLE IF EXISTS `dbinventory`.`rebates`;CREATE TABLE  `dbinventory`.`rebates` (`id` int(10) unsigned NOT NULL AUTO_INCREMENT,`customer_id` int(10) unsigned DEFAULT NULL,`total_rebate_amount` double(10,2) DEFAULT NULL,`total_qty_bought` double(10,2) DEFAULT NULL,`month` varchar(45) DEFAULT NULL,`issue_by` varchar(45) DEFAULT NULL,PRIMARY KEY (`id`) ) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;"



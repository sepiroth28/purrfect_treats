DROP TABLE IF EXISTS `dbinventory`.`rebate_price_table`;
CREATE TABLE  `dbinventory`.`rebate_price_table` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `qty_from` double(10,2) DEFAULT NULL,
  `qty_to` double(10,2) DEFAULT NULL,
  `applied_price` double(10,2) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;
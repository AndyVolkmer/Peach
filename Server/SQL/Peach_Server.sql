SET FOREIGN_KEY_CHECKS=0;
-- ----------------------------
-- Table structure for `accounts`
-- ----------------------------
DROP TABLE IF EXISTS `accounts`;
CREATE TABLE `accounts` (
  `ID` int(11) NOT NULL DEFAULT '0',
  `Name1` varchar(255) DEFAULT NULL,
  `Password1` varchar(255) DEFAULT NULL,
  `Time1` time DEFAULT NULL,
  `Date1` date DEFAULT NULL,
  `Banned1` varchar(255) DEFAULT NULL,
  `Level1` int(11) DEFAULT NULL,
  `SecretQuestion1` varchar(255) DEFAULT '',
  `SecretAnswer1` varchar(255) DEFAULT '',
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for `friends`
-- ----------------------------
DROP TABLE IF EXISTS `friends`;
CREATE TABLE `friends` (
  `ID` int(11) NOT NULL DEFAULT '0',
  `Name` varchar(20) NOT NULL DEFAULT '',
  `Friend` varchar(20) NOT NULL DEFAULT '',
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for `ignores`
-- ----------------------------
DROP TABLE IF EXISTS `ignores`;
CREATE TABLE `ignores` (
  `ID` int(11) NOT NULL DEFAULT '0',
  `Name` varchar(20) NOT NULL DEFAULT '',
  `IgnoredName` varchar(20) NOT NULL DEFAULT '',
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for `emotes`
-- ----------------------------
DROP TABLE IF EXISTS `emotes`;
CREATE TABLE `emotes` (
  `command` varchar(255) NOT NULL DEFAULT '',
  `is_user_text_1` varchar(255) DEFAULT '',
  `is_user_text_2` varchar(255) DEFAULT '',
  `is_not_user` varchar(255) DEFAULT '',
  `description` varchar(255) DEFAULT '',
  PRIMARY KEY (`command`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Table structure for `declinednames`
-- ----------------------------
DROP TABLE IF EXISTS `declinednames`;
CREATE TABLE `declinednames` (
  `Name` varchar(15) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;
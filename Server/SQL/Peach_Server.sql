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
  `Gender1` varchar(10) DEFAULT '',
  `Email1` varchar(255) DEFAULT '',
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- ----------------------------
-- Table structure for `commands`
-- ----------------------------
DROP TABLE IF EXISTS `commands`;
CREATE TABLE `commands` (
  `Syntax` varchar(255) NOT NULL DEFAULT '',
  `Description` varchar(255) DEFAULT '',
  PRIMARY KEY (`Syntax`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- ----------------------------
-- Records of commands
-- ----------------------------
INSERT INTO `commands` VALUES ('.accountinfo / .accinfo \'Account\'', 'Shows all information about that account.');
INSERT INTO `commands` VALUES ('.announce \'Text\'', 'Send a server side tagged message.');
INSERT INTO `commands` VALUES ('.ban \'Account\' [Reason]', 'Bans the account permanently.');
INSERT INTO `commands` VALUES ('.change gender \'Name\' \'Gender\'', 'Changes the \'Gender\' of \'Name\'.');
INSERT INTO `commands` VALUES ('.change level \'Name\' \'Level\'', 'Changes the \'Level\' of \'Name.');
INSERT INTO `commands` VALUES ('.change name \'Oldname\' \'Newname', 'Renames \'Oldname\' to \'Newname\'.');
INSERT INTO `commands` VALUES ('.change password \'Name\' \'Password\'', 'Changes the \'Password\' of \'Name\'.');
INSERT INTO `commands` VALUES ('.clear \'Name\'', 'Clears the chatbox of \'Name\'.');
INSERT INTO `commands` VALUES ('.delete \'Name\'', 'Deletes \'Name\'.');
INSERT INTO `commands` VALUES ('.gm [on / off]', 'Enables / Disables the gm flag in chat.');
INSERT INTO `commands` VALUES ('.help / .command', 'Shows a list of all avaible commands.');
INSERT INTO `commands` VALUES ('.kick \'Name\'', 'Disconnectes the user with \'Name\' from server.');
INSERT INTO `commands` VALUES ('.mute \'Name\' [Reason]', 'Mutes the user with \'Name\' permanently.');
INSERT INTO `commands` VALUES ('.reload \'Table\'', 'Reloads the \'Table\'.');
INSERT INTO `commands` VALUES ('.show accounts / onliners', 'Shows a list of all accounts / users avaible.');
INSERT INTO `commands` VALUES ('.unban account \'Account [Reason]', 'Unbans the account.');
INSERT INTO `commands` VALUES ('.unmute \'Name\' [Reason]', 'Removes the mute from \'Name\' if muted.');
INSERT INTO `commands` VALUES ('.userinfo \'Name\'', 'Shows a list with all information about \'Name\'.');

-- ----------------------------
-- Table structure for `declinednames`
-- ----------------------------
DROP TABLE IF EXISTS `declinednames`;
CREATE TABLE `declinednames` (
  `Name` varchar(255) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;


-- ----------------------------
-- Table structure for `emotes`
-- ----------------------------
DROP TABLE IF EXISTS `emotes`;
CREATE TABLE `emotes` (
  `command` varchar(255) NOT NULL DEFAULT '',
  `single_emote` varchar(255) NOT NULL,
  `target_emote` varchar(255) NOT NULL,
  PRIMARY KEY (`command`)
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
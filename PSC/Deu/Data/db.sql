--
-- Datenbank: `dbsv`
--

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_adm_logfile`
--

DROP TABLE IF EXISTS `dbsv_adm_logfile`;
CREATE TABLE IF NOT EXISTS `dbsv_adm_logfile` (
  `lid` smallint(5) NOT NULL AUTO_INCREMENT,
  `ltyp` varchar(1) NOT NULL,
  `laktion` varchar(255) NOT NULL,
  `ldatum` date NOT NULL,
  `luid` smallint(5) NOT NULL,
  PRIMARY KEY (`lid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=2 ;

--
-- Daten für Tabelle `dbsv_adm_logfile`
--

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_adm_roles`
--

DROP TABLE IF EXISTS `dbsv_adm_roles`;
CREATE TABLE IF NOT EXISTS `dbsv_adm_roles` (
  `rid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `rname` varchar(30) NOT NULL,
  `acl` varchar(4) NOT NULL,
  PRIMARY KEY (`rid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=4 ;

--
-- Daten für Tabelle `dbsv_adm_roles`
--

INSERT INTO `dbsv_adm_roles` VALUES(1, 'Administration', '1111');
INSERT INTO `dbsv_adm_roles` VALUES(2, 'Schueler', '0000');
INSERT INTO `dbsv_adm_roles` VALUES(3, 'Personal', '0111');

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_adm_user`
--

DROP TABLE IF EXISTS `dbsv_adm_user`;
CREATE TABLE IF NOT EXISTS `dbsv_adm_user` (
  `uid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `uname` varchar(20) NOT NULL,
  `upass` varchar(30) NOT NULL,
  `urole` smallint(5) unsigned NOT NULL,
  `udata` smallint(5) unsigned NOT NULL,
  PRIMARY KEY (`uid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=4 ;

--
-- Daten für Tabelle `dbsv_adm_user`
--

INSERT INTO `dbsv_adm_user` VALUES(1, 'admin', 'system', 1, 0);
INSERT INTO `dbsv_adm_user` VALUES(2, 'mueller', 'lehrer', 3, 0);
INSERT INTO `dbsv_adm_user` VALUES(3, 'test', 'nutzer', 2, 1);

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_main_class`
--

DROP TABLE IF EXISTS `dbsv_main_class`;
CREATE TABLE IF NOT EXISTS `dbsv_main_class` (
  `cid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `cname` varchar(25) NOT NULL,
  `descr` varchar(255) NULL,
  PRIMARY KEY (`cid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=4 ;

--
-- Daten für Tabelle `dbsv_main_class`
--

INSERT INTO `dbsv_main_class` VALUES(1, 'class 1', null);
INSERT INTO `dbsv_main_class` VALUES(2, 'class 2', null);
INSERT INTO `dbsv_main_class` VALUES(3, 'class 3', 'Lehrgang 1');

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_main_fach`
--

DROP TABLE IF EXISTS `dbsv_main_fach`;
CREATE TABLE IF NOT EXISTS `dbsv_main_fach` (
  `fid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `kname` varchar(10) NOT NULL,
  `name` varchar(50) NOT NULL,
  PRIMARY KEY (`fid`,`kname`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=5 ;

--
-- Daten für Tabelle `dbsv_main_fach`
--

INSERT INTO `dbsv_main_fach` VALUES(1, 'Eng', 'Englisch');
INSERT INTO `dbsv_main_fach` VALUES(2, 'Gesch', 'Geschichte');
INSERT INTO `dbsv_main_fach` VALUES(3, 'Ma', 'Mathematik');
INSERT INTO `dbsv_main_fach` VALUES(4, 'Basic', 'Computer Basics');

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_test_cat`
--

DROP TABLE IF EXISTS `dbsv_test_cat`;
CREATE TABLE IF NOT EXISTS `dbsv_test_cat` (
  `tcid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `tcorder` smallint(5) unsigned NOT NULL,
  `tctid` smallint(5) unsigned NOT NULL,
  `tctyp` tinyint(1) unsigned NOT NULL,
  `tcquest` varchar(500) NOT NULL,
  `tcans1` varchar(300) NOT NULL,
  `tcans2` varchar(300) NOT NULL,
  `tcans3` varchar(300) NOT NULL,
  `tcans4` varchar(300) NOT NULL,
  `tcans5` varchar(300) NOT NULL,
  `tcanswer` varchar(300) NOT NULL,
  `tcpoints` tinyint(2) unsigned NOT NULL,
  PRIMARY KEY (`tcid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;

--
-- Daten für Tabelle `dbsv_test_cat`
--

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_test_result`
--

DROP TABLE IF EXISTS `dbsv_test_result`;
CREATE TABLE IF NOT EXISTS `dbsv_test_result` (
  `trid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `truid` smallint(5) unsigned NOT NULL,
  `trtid` smallint(5) unsigned NOT NULL,
  `trfid` smallint(5) unsigned NOT NULL,
  `trdatum` varchar(20) NOT NULL,
  `trquestg` tinyint(3) unsigned NOT NULL,
  `trscore` smallint(4) unsigned NOT NULL,
  `trscore_max` smallint(4) unsigned NOT NULL,
  `trpass` tinyint(1) unsigned NOT NULL,
  PRIMARY KEY (`trid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;

--
-- Daten für Tabelle `dbsv_test_result`
--

-- --------------------------------------------------------

--
-- Tabellenstruktur für Tabelle `dbsv_test_setting`
--

DROP TABLE IF EXISTS `dbsv_test_setting`;
CREATE TABLE IF NOT EXISTS `dbsv_test_setting` (
  `tsid` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `tsuid` smallint(5) NOT NULL,
  `tsname` varchar(30) NOT NULL,
  `tsintro` varchar(500) NOT NULL,
  `tsclass` smallint(5) unsigned NOT NULL,
  `tsfach` smallint(5) NOT NULL,
  `tsactive` tinyint(1) unsigned NOT NULL,
  `tsallow_online` tinyint(1) unsigned NOT NULL,
  `tstimelimit` smallint(3) unsigned NOT NULL,
  `tstime_exp` tinyint(1) unsigned NOT NULL,
  `tsmultilimit` tinyint(2) unsigned NOT NULL,
  `tsdelay` smallint(4) unsigned NOT NULL,
  `tsshowq` tinyint(1) unsigned NOT NULL,
  `tsscore_pass` tinyint(3) unsigned NOT NULL,
  PRIMARY KEY (`tsid`)
) ENGINE=InnoDB  DEFAULT CHARSET=utf8 AUTO_INCREMENT=1 ;

--
-- Daten für Tabelle `dbsv_test_setting`
--

-- phpMyAdmin SQL Dump
-- version 4.8.3
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 08-01-2019 a las 10:24:40
-- Versión del servidor: 10.1.36-MariaDB
-- Versión de PHP: 7.2.11

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
SET AUTOCOMMIT = 0;
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `empresa`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `departaments`
--

CREATE TABLE `departaments` (
  `codi` smallint(5) UNSIGNED NOT NULL,
  `nom` varchar(30) DEFAULT NULL,
  `ciutat` varchar(20) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `departaments`
--

INSERT INTO `departaments` (`codi`, `nom`, `ciutat`) VALUES
(10, 'Comptabilitat', 'Barcelona'),
(20, 'Producció', 'Mollet del Vallès'),
(30, 'Comercial', 'Cardedeu'),
(40, 'Informàtica', 'Llinars del Vallès'),
(60, 'Salubritat', 'Barcelona'),
(70, 'Infermeria', 'Barcelona');

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `empleats`
--

CREATE TABLE `empleats` (
  `codi` int(11) NOT NULL,
  `nom` varchar(30) DEFAULT NULL,
  `funcio` varchar(20) DEFAULT NULL,
  `cap` int(11) DEFAULT NULL,
  `datacontracte` date DEFAULT NULL,
  `sou` int(11) DEFAULT NULL,
  `comisio` int(11) DEFAULT NULL,
  `ndepartament` smallint(5) UNSIGNED DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

--
-- Volcado de datos para la tabla `empleats`
--

INSERT INTO `empleats` (`codi`, `nom`, `funcio`, `cap`, `datacontracte`, `sou`, `comisio`, `ndepartament`) VALUES
(1465, 'Meseguer', 'Cap de departament', 7834, '2006-01-02', 2000, NULL, 40),
(1587, 'Mellado', 'Informàtic', 1465, '2007-03-01', 1400, NULL, 40),
(1647, 'Armengol', 'Informàtic', 1465, '2007-04-01', 1400, NULL, 40),
(1657, 'Garre', 'Informàtic', 1465, '2007-03-01', 1400, NULL, 40),
(1647, 'Mimbrero', 'Informàtic', 1465, '2007-04-01', 1400, NULL, 40),
(1657, 'Galán', 'Informàtic', 1465, '2007-04-10', 1400, NULL, 40),
(7369, 'Gamez', 'Administratiu', 7902, '0000-00-00', 1200, NULL, 20),
(7499, 'Antunez', 'Venedor', 7698, '2001-02-20', 1600, 300, 30),
(7521, 'Galán', 'Venedor', 7698, '0000-00-00', 1250, 500, 30),
(7566, 'Puig', 'Cap de departament', 7839, '0000-00-00', 2975, NULL, 20),
(7654, 'Martinez', 'Venedor', 7698, '2001-09-28', 1250, 1400, 30),
(7698, 'Duset', 'Cap de departament', 7839, '2001-05-01', 2850, NULL, 30),
(7782, 'Palmarola', 'Cap de departament', 7839, '2001-06-09', 2450, NULL, 10),
(7788, 'Segarra', 'Operari', 7566, '2005-07-21', 3000, NULL, 20),
(7839, 'Soto', 'Presidente', NULL, '2001-11-17', 5000, NULL, 10),
(7844, 'Llop', 'Venedor', 7698, '2001-09-08', 1500, NULL, 30),
(7876, 'Franquesa', 'Administratiu', 7788, '2005-08-24', 1100, NULL, 20),
(7900, 'Galvany', 'Administratiu', 7698, '2001-12-03', 950, NULL, 30),
(7902, 'Almiron', 'Operari', 7566, '2001-12-03', 3000, NULL, 20),
(7934, 'Saez', 'Administratiu', 7782, '2002-01-23', 1300, NULL, 10),
(8251, 'Luque', 'Administratiu', 7566, '2005-10-20', 1500, NULL, 20),
(7777, 'Valls', 'Subpresident', NULL, '0000-00-00', 4500, NULL, 40),
(60, 'Jordi Soto', 'Subpresident', NULL, '0000-00-00', 4500, NULL, NULL);

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `departaments`
--
ALTER TABLE `departaments`
  ADD PRIMARY KEY (`codi`);

--
-- Indices de la tabla `empleats`
--
ALTER TABLE `empleats`
  ADD KEY `cs_ndepartament` (`ndepartament`);

--
-- Restricciones para tablas volcadas
--

--
-- Filtros para la tabla `empleats`
--
ALTER TABLE `empleats`
  ADD CONSTRAINT `cs_ndepartament` FOREIGN KEY (`ndepartament`) REFERENCES `departaments` (`codi`);
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
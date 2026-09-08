-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Host: 127.0.0.1
-- Generation Time: Sep 08, 2026 at 02:02 AM
-- Server version: 10.4.32-MariaDB
-- PHP Version: 8.2.12

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Database: `db_eduboard`
--

-- --------------------------------------------------------

--
-- Table structure for table `storyboards`
--

CREATE TABLE `storyboards` (
  `id` int(11) NOT NULL,
  `username` varchar(100) DEFAULT NULL,
  `program_studi` varchar(255) DEFAULT NULL,
  `nama_mata_kuliah` varchar(255) DEFAULT NULL,
  `project_name` varchar(255) DEFAULT NULL,
  `learning_objectives` text DEFAULT NULL,
  `target_audience` varchar(100) DEFAULT NULL,
  `visual_style` varchar(100) DEFAULT NULL,
  `rps_filename` varchar(255) DEFAULT '-',
  `created_at` timestamp NOT NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `storyboards`
--

INSERT INTO `storyboards` (`id`, `username`, `program_studi`, `nama_mata_kuliah`, `project_name`, `learning_objectives`, `target_audience`, `visual_style`, `rps_filename`, `created_at`) VALUES
(8, 'Gerry', 'Magister Kecerdasan Buatan', 'Evaluasi dan Penerapan AI', 'Video Pembelajaran Pertemuan 1', 'Mata kuliah ini berfokus pada implementasi praktis dan tata kelola AI, bukan pada tuntutan agar mahasiswa menjadi programmer AI. Pendekatan ini menekankan kemampuan membangun prototipe, mengevaluasi kualitas output, melakukan audit, mengelola risiko, dan mengukur dampak organisasi.', 'General', 'Cinematic', 'Sistem_AI_Terapan_Evaluasi_dan_Tata_Kelola_Bahasa_Indonesia.pdf', '2026-09-07 07:01:38'),
(9, 'Gerry', 'Informatika', 'Sistem Basis Data', 'Video Pembelajaran Sistem Basis Data 1', 'Mahasiswa mampu mengimplementasikan pengelolaan basis data menggunakan SQL (DDL, DML, DCL) untuk membangun dan mengelola basis data sesuai kebutuhan aplikasi.', 'Mahasiswa S1/S2 (Advanced)', 'Cinematic', 'Sistem Basis Data.pdf', '2026-09-07 07:15:27'),
(10, 'Gerry', 'Informatika', 'Sistem Basis Data', 'Video Pembelajaran Sistem Basis Data 1', 'Mampu mengambil keputusan secara tepat dalam konteks penyelesaian masalah di bidang keahliannya, berdasarkan hasil analisis informasi dan data', 'Mahasiswa S1/S2 (Advanced)', 'Cinematic', 'Sistem Basis Data.pdf', '2026-09-07 07:16:42');

-- --------------------------------------------------------

--
-- Table structure for table `storyboard_scenes`
--

CREATE TABLE `storyboard_scenes` (
  `id` int(11) NOT NULL,
  `storyboard_id` int(11) DEFAULT NULL,
  `judul_scene` varchar(255) DEFAULT NULL,
  `visualisasi` text DEFAULT NULL,
  `instruksi_visual` text DEFAULT NULL,
  `animasi` text DEFAULT NULL,
  `on_screen_text` text DEFAULT NULL,
  `voice_over_text` text DEFAULT NULL,
  `backsound` varchar(100) DEFAULT NULL,
  `durasi` varchar(50) DEFAULT NULL,
  `saran_info` text DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `storyboard_scenes`
--

INSERT INTO `storyboard_scenes` (`id`, `storyboard_id`, `judul_scene`, `visualisasi`, `instruksi_visual`, `animasi`, `on_screen_text`, `voice_over_text`, `backsound`, `durasi`, `saran_info`) VALUES
(64, 8, '01 - Pengenalan Konsep', 'Ilustrasi diagram blok', 'Zoom in ke bagian matriks', 'Fade in / slide from left', 'Rumus Utama CNN', 'Pernahkah kalian berpikir bahwa AI dapat membantu organisasi dalam membuat keputusan?', 'Ambient Tech - Calm', '00:15', 'Pastikan visual matriks jelas'),
(65, 8, '02 - Memahami AI', 'Grafik flowchart', 'Klik untuk melihat bagian lain', 'Slide from top', 'Definisi AI dan Machine Learning', 'AI adalah teknologi yang dapat mempelajari dan menirukan perilaku manusia', 'Soft Piano - Melodic', '00:30', 'Berikan waktu untuk memahami definisi AI'),
(66, 8, '03 - Membangun Prototype', 'Ilustrasi diagram workflow', 'Klik untuk memperbesar', 'Zoom in', 'Membangun Prototype AI', 'Prototype AI adalah tahapan awal dalam pengembangan AI', 'Nature Sounds - Birdsong', '00:45', 'Perhatikan detail penggunaan platform low-code/no-code'),
(67, 8, '04 - Menguji AI', 'Grafik bar chart', 'Klik untuk membandingkan', 'Slide from right', 'Evaluasi Kualitas Output AI', 'Evaluasi kualitas output AI dapat membantu dalam menentukan keberhasilan AI', 'Electronic Music - Synth', '01:00', 'Perhatikan keterbatasan evaluasi kualitas output AI'),
(68, 8, '05 - Mengaudit AI', 'Ilustrasi diagram audit', 'Klik untuk memperbesar', 'Fade in', 'Audit Keluaran Model AI', 'Audit keluaran model AI dapat membantu dalam menentukan keberhasilan AI', 'Soft Rock - Guitar', '01:15', 'Perhatikan keterbatasan audit keluaran model AI'),
(69, 8, '06 - Mengelola Risiko', 'Grafik pie chart', 'Klik untuk membandingkan', 'Slide from bottom', 'Mengelola Risiko AI', 'Mengelola risiko AI dapat membantu dalam menentukan keberhasilan AI', 'Nature Sounds - Waterfall', '01:30', 'Perhatikan keterbatasan mengelola risiko AI'),
(70, 8, '07 - Memastikan Kepatuhan', 'Ilustrasi diagram compliance', 'Klik untuk memperbesar', 'Fade out', 'Memastikan Kepatuhan AI', 'Memastikan kepatuhan AI dapat membantu dalam menentukan keberhasilan AI', 'Electronic Music - Drum', '01:45', 'Perhatikan keterbatasan memastikan kepatuhan AI'),
(71, 8, '08 - Mengukur Dampak Bisnis', 'Grafik line chart', 'Klik untuk membandingkan', 'Slide from top', 'Mengukur Dampak Bisnis AI', 'Mengukur dampak bisnis AI dapat membantu dalam menentukan keberhasilan AI', 'Soft Rock - Bass', '02:00', 'Perhatikan keterbatasan mengukur dampak bisnis AI'),
(72, 8, '09 - Sesi 13 - Monitoring Output Drift & Penurunan Kualitas', 'Ilustrasi diagram drift', 'Klik untuk memperbesar', 'Fade in', 'Monitoring Output Drift & Penurunan Kualitas AI', 'Monitoring output drift dan penurunan kualitas AI dapat membantu dalam menentukan keberhasilan AI', 'Nature Sounds - Birdsong', '02:15', 'Perhatikan keterbatasan monitoring output drift dan penurunan kualitas AI'),
(73, 8, '10 - Sesi 14 - Final Capstone Defense: Prototipe AI Praktis & Audit', 'Ilustrasi diagram workflow', 'Klik untuk memperbesar', 'Slide from bottom', 'Final Capstone Defense: Prototipe AI Praktis & Audit', 'Final capstone defense: prototipe AI praktis & audit dapat membantu dalam menentukan keberhasilan AI', 'Electronic Music - Synth', '02:30', 'Perhatikan keterbatasan final capstone defense: prototipe AI praktis & audit'),
(74, 9, '01 - Pengenalan Konsep', 'Ilustrasi diagram blok Sistem Basis Data', 'Zoom in ke bagian matriks', 'Fade in / slide from left', 'Konsep dasar Sistem Basis Data', 'Pernahkah kalian berpikir tentang bagaimana data disimpan dan dikelola?', 'Ambient Tech - Calm', '00:10', 'Pastikan visual matriks jelas'),
(75, 9, '02 - Aksesori DBMS', 'Ilustrasi diagram blok DBMS', 'Zoom in ke bagian SQL', 'Fade in / slide from left', 'Aksesori DBMS', 'Apa itu Structured Query Language (SQL)?', 'Ambient Tech - Calm', '00:12', 'Perhatikan sintaks SQL'),
(76, 9, '03 - Model Data', 'Ilustrasi diagram blok Model Data', 'Zoom in ke bagian Model Relasional', 'Fade in / slide from left', 'Model Data', 'Apa itu Model Data dan bagaimana keterkaitan antara Model Data dengan Model Relasional?', 'Ambient Tech - Calm', '00:15', 'Perhatikan keterkaitan antara Model Data dengan Model Relasional'),
(77, 9, '04 - Pengelolaan Basis Data', 'Ilustrasi diagram blok Pengelolaan Basis Data', 'Zoom in ke bagian SQL DDL dan DML', 'Fade in / slide from left', 'Pengelolaan Basis Data', 'Bagaimana kalian mengelola basis data menggunakan SQL (DDL, DML, DCL)?', 'Ambient Tech - Calm', '00:18', 'Perhatikan sintaks SQL DDL dan DML'),
(78, 9, '05 - Manajemen Transaksi', 'Ilustrasi diagram blok Manajemen Transaksi', 'Zoom in ke bagian Manajemen Transaksi', 'Fade in / slide from left', 'Manajemen Transaksi', 'Bagaimana kalian menjaga konsistensi, integritas, keamanan, dan konkurensi data dalam penyelesaian studi kasus?', 'Ambient Tech - Calm', '00:20', 'Perhatikan pentingnya Manajemen Transaksi dalam penyelesaian studi kasus'),
(79, 10, '01 - Pengenalan Konsep Sistem Basis Data', 'Diagram bloklarge database systems', 'Zoom in to database schema', 'Fade in / slide from left', 'Database Systems: Definition and Types', 'In today\'s lesson, we will explore the concept of database systems and their importance in computer science.', 'Ambient Tech - Calm', '00:30', 'Make sure to understand the concept of database systems'),
(80, 10, '02 - Rancangan Basis Data', 'ERD diagram with entities and relationships', 'Click to zoom in on entity relationships', 'Fade in / slide from right', 'Entity-Relationship Diagram (ERD): Designing Database Structures', 'In this scene, we will introduce the concept of ERD and its importance in database design.', 'Techno - Focused', '00:45', 'Practice designing ERD for a given scenario'),
(81, 10, '03 - Normalisasi', 'Table normalization with 1NF, 2NF, and 3NF', 'Click to see the normalization process', 'Fade in / slide from bottom', 'Database Normalization: 1NF, 2NF, and 3NF', 'In this scene, we will explore the concept of database normalization and its importance in database design.', 'Calm Music - Piano', '01:00', 'Understand the concept of normalization and its applications'),
(82, 10, '04 - Implementasi SQL', 'SQL query with SELECT, INSERT, UPDATE, and DELETE statements', 'Click to execute the SQL query', 'Fade in / slide from top', 'SQL: SELECT, INSERT, UPDATE, and DELETE Statements', 'In this scene, we will introduce the concept of SQL and its importance in database management.', 'Techno - Fast-paced', '01:15', 'Practice writing SQL queries for a given scenario'),
(83, 10, '05 - Presentasi Solusi', 'Presentasi with slides and data visualization', 'Click to navigate through the slides', 'Fade in / slide from left', 'Presenting Solutions: Data Visualization and Storytelling', 'In this scene, we will introduce the concept of presenting solutions and its importance in database management.', 'Calm Music - Nature', '01:30', 'Practice presenting solutions for a given scenario');

-- --------------------------------------------------------

--
-- Table structure for table `users`
--

CREATE TABLE `users` (
  `id` int(11) NOT NULL,
  `username` varchar(100) NOT NULL,
  `password` varchar(255) NOT NULL,
  `created_at` timestamp NOT NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Dumping data for table `users`
--

INSERT INTO `users` (`id`, `username`, `password`, `created_at`) VALUES
(1, 'Gerry', 'gerry123', '2026-08-24 02:29:53'),
(2, 'admin', 'admin', '2026-08-24 02:51:55');

--
-- Indexes for dumped tables
--

--
-- Indexes for table `storyboards`
--
ALTER TABLE `storyboards`
  ADD PRIMARY KEY (`id`);

--
-- Indexes for table `storyboard_scenes`
--
ALTER TABLE `storyboard_scenes`
  ADD PRIMARY KEY (`id`),
  ADD KEY `storyboard_id` (`storyboard_id`);

--
-- Indexes for table `users`
--
ALTER TABLE `users`
  ADD PRIMARY KEY (`id`),
  ADD UNIQUE KEY `username` (`username`);

--
-- AUTO_INCREMENT for dumped tables
--

--
-- AUTO_INCREMENT for table `storyboards`
--
ALTER TABLE `storyboards`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=11;

--
-- AUTO_INCREMENT for table `storyboard_scenes`
--
ALTER TABLE `storyboard_scenes`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=84;

--
-- AUTO_INCREMENT for table `users`
--
ALTER TABLE `users`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=3;

--
-- Constraints for dumped tables
--

--
-- Constraints for table `storyboard_scenes`
--
ALTER TABLE `storyboard_scenes`
  ADD CONSTRAINT `storyboard_scenes_ibfk_1` FOREIGN KEY (`storyboard_id`) REFERENCES `storyboards` (`id`) ON DELETE CASCADE;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;

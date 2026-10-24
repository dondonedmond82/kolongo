-- SOCASP / Fuel data lookup tables
-- Safe to run repeatedly: unique names + ON DUPLICATE KEY UPDATE prevent duplicates.
-- Run this script in the MySQL database used by your application.

CREATE TABLE IF NOT EXISTS origine (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    nom VARCHAR(100) NOT NULL,
    PRIMARY KEY (id),
    UNIQUE KEY uq_origine_nom (nom)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS marketeur (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    nom VARCHAR(150) NOT NULL,
    PRIMARY KEY (id),
    UNIQUE KEY uq_marketeur_nom (nom)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS provenance (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    ville VARCHAR(150) NOT NULL,
    pays VARCHAR(150) NULL,
    PRIMARY KEY (id),
    UNIQUE KEY uq_provenance_ville_pays (ville, pays)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

INSERT INTO origine (nom) VALUES
    ('Air'),
    ('Terrestre'),
    ('Fleuve')
ON DUPLICATE KEY UPDATE nom = VALUES(nom);

INSERT INTO marketeur (nom) VALUES
    ('Total'),
    ('Tradex'),
    ('Elf'),
    ('Green Oil'),
    ('Neptune')
ON DUPLICATE KEY UPDATE nom = VALUES(nom);

-- Cities supplied in provenance_data.xlsx.
-- Country is left NULL because the spreadsheet does not specify each city's country.
INSERT INTO provenance (ville, pays) VALUES
    ('Douala', NULL),
    ('Brazzaville', NULL),
    ('Limbe', NULL),
    ('Kinshasa', NULL)
ON DUPLICATE KEY UPDATE ville = VALUES(ville);

-- Example: use IDs in your main fuel/importation/distribution table rather than repeating text.
-- SELECT id, nom FROM origine ORDER BY nom;
-- SELECT id, nom FROM marketeur ORDER BY nom;
-- SELECT id, ville, pays FROM provenance ORDER BY ville;

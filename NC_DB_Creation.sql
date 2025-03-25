-- Création de la base de données NC_DB
CREATE DATABASE NC_DB;
GO

-- Utilisation de la base de données NC_DB
USE NC_DB;
GO

-- Création de la table Users
CREATE TABLE Users (
    ID_User INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    ID_Citrix VARCHAR(50) NOT NULL,
    Login VARCHAR(50) NOT NULL,
    Password VARCHAR(100) NOT NULL,
    Role VARCHAR(50) NOT NULL,
    CONSTRAINT UK_Users_ID_Citrix UNIQUE (ID_Citrix)
);
GO

-- Création de la table Effectif
CREATE TABLE Effectif (
    ID_Effectif INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    ID_Citrix VARCHAR(50) NOT NULL,
    Nom VARCHAR(50) NOT NULL,
    Prenom VARCHAR(50) NOT NULL,
    Tel VARCHAR(20),
    Adresse VARCHAR(100),
    DateIn DATE,
    DateOut DATE,
    Statut VARCHAR(50),
    CONSTRAINT FK_Effectif_Users FOREIGN KEY (ID_Citrix) REFERENCES Users(ID_Citrix)
);
GO

-- Création de la table Fiches
CREATE TABLE Fiches (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Num_Fiche VARCHAR(50) NOT NULL,
    Date_Fiche DATE NOT NULL,
    Ref VARCHAR(50),
    Source VARCHAR(100),
    Motif VARCHAR(255),
    Priorite VARCHAR(20),
    Statut_Fiche VARCHAR(50),
    Commentaire TEXT,
    CONSTRAINT UK_Fiches_Num_Fiche UNIQUE (Num_Fiche)
);
GO

-- Création de la table Action_Fiches
CREATE TABLE Action_Fiches (
    ID INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    Num_Fiche VARCHAR(50) NOT NULL,
    Date_Action DATETIME NOT NULL,
    Source VARCHAR(100),
    Motif_Action VARCHAR(255),
    Action_Fiche TEXT,
    Priorite VARCHAR(20),
    Statut_Action VARCHAR(50),
    Resume_Action_Fiche TEXT,
    Commentaire_Action_Fiche TEXT,
    Date_Rappel DATE,
    Creneau_Rappel VARCHAR(50),
    Statut_Rappel VARCHAR(50),
    CONSTRAINT FK_Action_Fiches_Fiches FOREIGN KEY (Num_Fiche) REFERENCES Fiches(Num_Fiche)
);
GO

-- Ajout d'index pour améliorer les performances
CREATE INDEX IX_Users_ID_Citrix ON Users(ID_Citrix);
CREATE INDEX IX_Effectif_ID_Citrix ON Effectif(ID_Citrix);
CREATE INDEX IX_Fiches_Num_Fiche ON Fiches(Num_Fiche);
CREATE INDEX IX_Action_Fiches_Num_Fiche ON Action_Fiches(Num_Fiche);
GO

-- Insertion de données de test (optionnel)
INSERT INTO Users (ID_Citrix, Login, Password, Role)
VALUES 
('S800100', 'admin', 'motdepasse', 'Administrateur'),
('S800101', 'user1', 'password1', 'Utilisateur'),
('S800102', 'user2', 'password2', 'Technicien');
GO
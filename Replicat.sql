-- Création de la base de données NC_DB_REPLICA
CREATE DATABASE NC_DB_REPLICA;
GO

-- Utilisation de la base de données NC_DB_REPLICA
USE NC_DB_REPLICA;
GO

-- Création de la table Fiches (version répliquée)
CREATE TABLE Fiches (
    ID INT NOT NULL PRIMARY KEY,  -- Suppression de IDENTITY pour la réplication
    Num_Fiche VARCHAR(50) NOT NULL,
    Date_Fiche DATE NOT NULL,
    Ref VARCHAR(50),
    Source VARCHAR(100),
    Motif VARCHAR(255),
    Priorite VARCHAR(20),
    Statut_Fiche VARCHAR(50),
    Commentaire TEXT,
    LastUpdated DATETIME DEFAULT GETDATE(),  -- Ajout pour le suivi des modifications
    CONSTRAINT UK_Fiches_Num_Fiche UNIQUE (Num_Fiche)
);
GO

-- Création de la table Action_Fiches (version répliquée)
CREATE TABLE Action_Fiches (
    ID INT NOT NULL PRIMARY KEY,  -- Suppression de IDENTITY pour la réplication
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
    LastUpdated DATETIME DEFAULT GETDATE(),  -- Ajout pour le suivi des modifications
    CONSTRAINT FK_Action_Fiches_Fiches FOREIGN KEY (Num_Fiche) REFERENCES Fiches(Num_Fiche)
);
GO

-- Ajout d'index pour améliorer les performances
CREATE INDEX IX_Fiches_Num_Fiche ON Fiches(Num_Fiche);
CREATE INDEX IX_Action_Fiches_Num_Fiche ON Action_Fiches(Num_Fiche);
CREATE INDEX IX_Fiches_LastUpdated ON Fiches(LastUpdated);
CREATE INDEX IX_Action_Fiches_LastUpdated ON Action_Fiches(LastUpdated);
GO

-- Procédure stockée pour synchroniser les données
CREATE PROCEDURE usp_SyncFichesData
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Synchronisation des fiches
    MERGE INTO NC_DB_REPLICA.dbo.Fiches AS target
    USING NC_DB.dbo.Fiches AS source
    ON target.ID = source.ID
    WHEN MATCHED AND target.LastUpdated < source.LastUpdated THEN
        UPDATE SET 
            target.Num_Fiche = source.Num_Fiche,
            target.Date_Fiche = source.Date_Fiche,
            target.Ref = source.Ref,
            target.Source = source.Source,
            target.Motif = source.Motif,
            target.Priorite = source.Priorite,
            target.Statut_Fiche = source.Statut_Fiche,
            target.Commentaire = source.Commentaire,
            target.LastUpdated = GETDATE()
    WHEN NOT MATCHED BY TARGET THEN
        INSERT (ID, Num_Fiche, Date_Fiche, Ref, Source, Motif, Priorite, Statut_Fiche, Commentaire, LastUpdated)
        VALUES (source.ID, source.Num_Fiche, source.Date_Fiche, source.Ref, source.Source, source.Motif, 
                source.Priorite, source.Statut_Fiche, source.Commentaire, GETDATE());
    
    -- Synchronisation des actions fiches
    MERGE INTO NC_DB_REPLICA.dbo.Action_Fiches AS target
    USING NC_DB.dbo.Action_Fiches AS source
    ON target.ID = source.ID
    WHEN MATCHED AND target.LastUpdated < source.LastUpdated THEN
        UPDATE SET 
            target.Num_Fiche = source.Num_Fiche,
            target.Date_Action = source.Date_Action,
            target.Source = source.Source,
            target.Motif_Action = source.Motif_Action,
            target.Action_Fiche = source.Action_Fiche,
            target.Priorite = source.Priorite,
            target.Statut_Action = source.Statut_Action,
            target.Resume_Action_Fiche = source.Resume_Action_Fiche,
            target.Commentaire_Action_Fiche = source.Commentaire_Action_Fiche,
            target.Date_Rappel = source.Date_Rappel,
            target.Creneau_Rappel = source.Creneau_Rappel,
            target.Statut_Rappel = source.Statut_Rappel,
            target.LastUpdated = GETDATE()
    WHEN NOT MATCHED BY TARGET THEN
        INSERT (ID, Num_Fiche, Date_Action, Source, Motif_Action, Action_Fiche, Priorite, 
                Statut_Action, Resume_Action_Fiche, Commentaire_Action_Fiche, 
                Date_Rappel, Creneau_Rappel, Statut_Rappel, LastUpdated)
        VALUES (source.ID, source.Num_Fiche, source.Date_Action, source.Source, source.Motif_Action, 
                source.Action_Fiche, source.Priorite, source.Statut_Action, source.Resume_Action_Fiche, 
                source.Commentaire_Action_Fiche, source.Date_Rappel, source.Creneau_Rappel, 
                source.Statut_Rappel, GETDATE());
END;
GO

-- Déclencheur pour mettre à jour LastUpdated
CREATE TRIGGER trg_Fiches_Update
ON Fiches
AFTER UPDATE
AS
BEGIN
    UPDATE Fiches
    SET LastUpdated = GETDATE()
    FROM Fiches f
    INNER JOIN inserted i ON f.ID = i.ID
END;
GO

CREATE TRIGGER trg_ActionFiches_Update
ON Action_Fiches
AFTER UPDATE
AS
BEGIN
    UPDATE Action_Fiches
    SET LastUpdated = GETDATE()
    FROM Action_Fiches af
    INNER JOIN inserted i ON af.ID = i.ID
END;
GO
#!/usr/bin/env python
# coding: utf-8

"""
Ce script Python analyse et visualise les données de performance d'une machine.
Il utilise les bibliothèques pandas, matplotlib et numpy.

"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from math import sqrt, pow
from random import randint

# Définir le chemin vers les données (à adapter si nécessaire)
DATA_DIR = "refinedData"
OUTPUT_DIR = "output"  # Répertoire pour sauvegarder les graphiques

def load_data(data_dir):
    """
    Charge les données à partir des fichiers Excel.

    Args:
        data_dir (str): Le chemin vers le répertoire contenant les fichiers Excel.

    Returns:
        tuple: Un tuple contenant les DataFrames df, df2 et df3.
               Retourne None si une erreur se produit lors du chargement des données.
    """
    try:
        # Lecture des fichiers Excel avec gestion des exceptions pour les erreurs de fichier
        df = pd.read_excel(os.path.join(os.getcwd(), data_dir, "Performance.xlsx"), sheet_name="Feuil1")
        df2 = pd.read_excel(os.path.join(os.getcwd(), data_dir, "Performance.xlsx"), sheet_name="Feuil4")
        df3 = pd.read_excel(os.path.join(os.getcwd(), data_dir, "Performance.xlsx"), sheet_name="Feuil5")

        # Définition de l'index et renommage des colonnes
        df.set_index(df["Unnamed: 0"], inplace=True)
        df.columns = range(28)  # Utilisation de range pour renommer les colonnes (plus pythonique)

        return df, df2, df3

    except FileNotFoundError:
        # Gestion spécifique de l'erreur si les fichiers ne sont pas trouvés
        print(f"Erreur: Les fichiers de données n'ont pas été trouvés dans '{data_dir}'.")
        return None, None, None
    except Exception as e:
        # Gestion des autres exceptions possibles lors de la lecture des fichiers
        print(f"Une erreur inattendue s'est produite lors du chargement des données : {e}")
        return None, None, None

def calculate_performance_metrics(df):
    """
    Calcule les indicateurs clés de performance à partir du DataFrame df.

    Args:
        df (pd.DataFrame): Le DataFrame contenant les données de performance.

    Returns:
        dict: Un dictionnaire contenant les indicateurs de performance calculés.
              Retourne None si le DataFrame d'entrée est None ou vide.
    """

    # Vérification de la validité du DataFrame en entrée
    if df is None or df.empty:
        print("Erreur: DataFrame invalide. Impossible de calculer les métriques de performance.")
        return None

    # Définition des jours de début et de fin de la période d'analyse
    start_day = df.columns[1]
    end_day = df.columns[26]

    # Extraction des données de performance et stockage dans un dictionnaire
    performance_data = {
        "disponibilite": df.loc["Disponibilité %", start_day:end_day],
        "heures_planifiees": df.loc["Heures Planifiés", start_day:end_day],
        "heures_marche": df.loc["Heures de Marche", start_day:end_day],
        "perf_planifiee_ml_min": df.loc["mL/min Planifiées", start_day:end_day],
        "perf_marche_ml_min": df.loc["mL/min Marche", start_day:end_day],
        "perf_marche_m2_h": df.loc["000m²/hrs Marche", start_day:end_day],
        "perf_planifiee_m2_h": df.loc["000m²/Hrs Planifiées", start_day:end_day],
        "vitesse_cible_m2_h": df.loc["Vitesse Cible (000m²/hrs Planifiées)", start_day:end_day] * 1000,
        "% Arrêts Techniques" : df.loc["% Arrêts Techniques",start_day:end_day],
        "% Arrêts Opérationnels" : df.loc["% Arrêts Opérationnels",start_day:end_day],
        "MTBF" : df.loc["MTBF", start_day:end_day],
        "MTTR" : df.loc["MTTR",start_day:end_day],
        "Surface Fabriquée" : df.loc["Surface Fabriquée",start_day:end_day],
        "mL Fabriquées" : df.loc["mL Fabriquées",start_day:end_day]
    }

    # Calculs des indicateurs dérivés
    performance_data["arrets"] = performance_data["heures_planifiees"] - performance_data["heures_marche"] # Calcul des heures d'arrêt
    performance_data["ecart_vitesse"] = performance_data["vitesse_cible_m2_h"] - performance_data["perf_marche_m2_h"] # Calcul de l'écart de vitesse
    performance_data["temps_arret_tech"] = performance_data["heures_planifiees"] * performance_data["% Arrêts Techniques"] # Calcul du temps d'arrêt technique
    performance_data["temps_arret_op"] = performance_data["heures_planifiees"] * performance_data["% Arrêts Opérationnels"] # Calcul du temps d'arrêt opérationnel
    performance_data["prod_cible_h_marche"] = performance_data["heures_marche"] * 210 * 60 # Calcul de la production cible basée sur les heures de marche
    performance_data["prod_theorique"] = performance_data["heures_planifiees"] * 210 * 60 # Calcul de la production théorique
    performance_data["ecart_prod"] = performance_data["prod_theorique"] - performance_data["mL Fabriquées"] # Calcul de l'écart de production
    performance_data["perte_prod_tech"] = performance_data["temps_arret_tech"] * 210 * 60 # Calcul de la perte de production due aux arrêts techniques
    performance_data["perte_prod_op"] = performance_data["temps_arret_op"] * 210 * 60 # Calcul de la perte de production due aux arrêts opérationnels

    return performance_data

def create_summary_dataframes(performance_data):
    """
    Crée des DataFrames récapitulatifs pour l'analyse.

    Args:
        performance_data (dict): Le dictionnaire contenant les données de performance calculées.

    Returns:
        tuple: Un tuple contenant les DataFrames v_cible, cause et resultat.
               Retourne None, None, None si performance_data est None.
    """

    # Vérification de la validité des données en entrée
    if performance_data is None:
        return None, None, None

    # DataFrame pour la vitesse cible
    v_cible = pd.DataFrame({
        "jour": list(str(jour) for jour in performance_data["heures_planifiees"].columns), # Création de la colonne 'jour'
        "vitesse_cible": [185 for _ in range(len(performance_data["heures_planifiees"].columns))] # Création de la colonne 'vitesse_cible'
    })

    # DataFrame pour l'analyse des causes de perte de production
    total_ecart_prod = performance_data["ecart_prod"].sum() # Calcul de l'écart total de production
    perte_prod_tech_totale = performance_data["perte_prod_tech"].sum() # Calcul de la perte de production totale due aux arrêts techniques
    perte_prod_op_totale = performance_data["perte_prod_op"].sum() # Calcul de la perte de production totale due aux arrêts opérationnels

    cause = pd.DataFrame([{
        "arrêts_technique": perte_prod_tech_totale / total_ecart_prod if total_ecart_prod else 0, # Calcul de la contribution des arrêts techniques
        "vitesse_machine": 1 - (perte_prod_tech_totale + perte_prod_op_totale) / total_ecart_prod if total_ecart_prod else 0, # Calcul de la contribution de la vitesse de la machine
        "arrêts_opérationnels": perte_prod_op_totale / total_ecart_prod if total_ecart_prod else 0 # Calcul de la contribution des arrêts opérationnels
    }])

    # DataFrame pour l'amélioration de la productivité
    resultat_data = {
        "jour": list(str(jour) for jour in performance_data["heures_planifiees"].columns), # Création de la colonne 'jour'
        "amelioration": [ # Calcul de l'amélioration de la productivité pour chaque jour
            (h_marche + 0.2 * arret) * 60 * perf_m / (h_planifie * 60)
            if perf_m >= 210 * 0.95
            else (h_marche + 0.2 * arret) * 60 * 210 * 0.95 / (h_planifie * 60)
            for h_marche, arret, perf_m, h_planifie in zip(
                performance_data["heures_marche"],
                performance_data["arrets"],
                performance_data["perf_marche_ml_min"],
                performance_data["heures_planifiees"],
            )
        ]
    }
    resultat = pd.DataFrame(resultat_data) # Création du DataFrame
    resultat.set_index("jour", inplace=True) # Définition de la colonne 'jour' comme index

    return v_cible, cause, resultat

def calculate_trs(performance_data, qualite_df):
    """
    Calculer le TRS (Taux de Rendement Synthétique).

    Args:
        performance_data (dict): Dictionnaire contenant les données de performance.
        qualite_df (pd.DataFrame): DataFrame contenant les données de qualité.

    Returns:
        tuple: TRS moyen et DataFrame des TRS quotidiens.
               Retourne None, None si performance_data ou qualite_df sont None.
    """
    # Vérification de la validité des données en entrée
    if performance_data is None or qualite_df is None:
        return None, None

    # Calcul des composantes du TRS
    m_perf = performance_data["perf_marche_ml_min"].mean() # Performance moyenne
    m_dispo = performance_data["disponibilite"].mean() # Disponibilité moyenne
    m_qualite = 1 - qualite_df.mean()  # Qualité moyenne (en supposant que qualite_df représente les défauts)

    # Calcul du TRS moyen et quotidien
    trs_moyen = (m_perf / 185) * m_dispo * m_qualite
    trs_quotidien = (performance_data["perf_marche_ml_min"] / 185) * performance_data["disponibilite"] * (1 - qualite_df)

    return trs_moyen, trs_quotidien

def calculate_benefices(performance_data, qualite_df):
    """
    Calculer les bénéfices actuels et après améliorations.

    Args:
        performance_data (dict): Dictionnaire contenant les données de performance.
        qualite_df (pd.DataFrame): DataFrame contenant les données de qualité.

    Returns:
        tuple: Liste des bénéfices actuels et liste des bénéfices après améliorations.
               Retourne None, None si performance_data ou qualite_df sont None.
    """
    # Vérification de la validité des données en entrée
    if performance_data is None or qualite_df is None:
        return None, None

    # Calcul des bénéfices actuels
    benefices_actuels = [a * (2.5 - b * 7.5) for a, b in zip(performance_data["Surface Fabriquée"], qualite_df)]
    # Calcul des bénéfices après améliorations (formule complexe, nécessite une compréhension approfondie du contexte)
    benefices_apres = [
        2.  5 * (0.2 * a * b + 0.95 * c * 27770) - 7.5 * f * c * 27770
        if b < 0.95 * 27770
        else 2.5 * (0.2 * a * b + b * c) - 7.5 * f * c * 27770
        for a, b, c, f in zip(
            performance_data["arrets"],
            performance_data["perf_marche_m2_h"],
            performance_data["heures_marche"],
            qualite_df,
        )
    ]

    return benefices_actuels, benefices_apres


def plot_performance_graphs(performance_data, v_cible, resultat, trs_quotidien, cause, qualite, benefices_actuels, benefices_apres):
    """
    Génère et sauvegarde les graphiques de performance.

    Args:
        performance_data (dict): Dictionnaire contenant les données de performance.
        v_cible (pd.DataFrame): DataFrame contenant les données de vitesse cible.
        resultat (pd.DataFrame): DataFrame contenant les données d'amélioration.
        trs_quotidien (pd.DataFrame): DataFrame contenant les données de TRS quotidien.
        cause (pd.DataFrame): DataFrame contenant les causes de perte de production.
        qualite (pd.Series): Series contenant les données de qualité.
        benefices_actuels (list): Liste des bénéfices actuels.
        benefices_apres (list): Liste des bénéfices après améliorations.
    """

    # Assurez-vous que le répertoire de sortie existe (créez-le s'il n'existe pas)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    jours = performance_data["heures_planifiees"].columns  # Récupère les noms des jours depuis les données
    jours_numeric = range(1, len(jours) + 1)  # Crée une liste numérique pour l'axe X (si nécessaire)

    # 1. Graphique du TRS (Taux de Rendement Synthétique)
    plt.figure(figsize=(10, 5))  # Crée une nouvelle figure pour le graphique (taille ajustable)
    plt.plot(jours, trs_quotidien, marker='o', color="#3498db")  # Trace le TRS en fonction des jours
    plt.title("Taux de rendement synthétique durant le mois de février 2025", fontsize=14)  # Titre du graphique
    plt.xlabel("jours", color="blue")  # Étiquette de l'axe X
    plt.ylabel("TRS", color="blue")  # Étiquette de l'axe Y
    plt.grid(True, which="both", alpha=0.6)  # Ajoute une grille au graphique (lignes principales et secondaires, transparence)
    plt.xticks(rotation=45, ha='right')  # Incline les étiquettes de l'axe X pour une meilleure lisibilité
    plt.tight_layout()  # Ajuste automatiquement la mise en page pour éviter les chevauchements
    plt.savefig(os.path.join(OUTPUT_DIR, "TRS.png"))  # Sauvegarde le graphique dans un fichier
    plt.show()  # Affiche le graphique à l'écran

    # 2. Graphique de la disponibilité
    plt.figure(figsize=(10, 5))
    plt.plot(jours, performance_data["disponibilite"], marker='o', color="#2ecc71")
    plt.title("Fluctuation du taux de disponibilité de la machine Onduleuse", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("Disponibilité de l'Onduleuse", color="blue")
    plt.grid(True, which="both", alpha=0.6)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "disponibilite.png"))
    plt.show()

    # 3. Graphique des heures planifiées vs heures de marche
    plt.figure(figsize=(10, 5))
    plt.plot(jours, performance_data["heures_planifiees"], color="green", marker='o', label="Heures Planifiées")  # Trace les heures planifiées
    plt.plot(jours, performance_data["heures_marche"], color="red", marker='x', label="Heures de Marche")  # Trace les heures de marche
    plt.title("Heures planifié vs Heures de marche mois Février", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("Heures", color="blue")
    plt.grid(True, which="both", alpha=0.6)
    plt.xticks(rotation=45, ha='right')
    plt.legend()  # Affiche la légende pour distinguer les séries de données
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "heures_planifiees_marche.png"))
    plt.show()

    # 4. Graphique des heures d'arrêt
    plt.figure(figsize=(10, 5))
    plt.bar(jours, performance_data["arrets"], color="orange")  # Crée un graphique à barres pour les heures d'arrêt
    plt.title("Ecart entre les heures planifiées et les heures de marche mois février", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("Heures d'arrêt", color="blue")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "heures_arret.png"))
    plt.show()

    # 5. Graphique de la productivité
    plt.figure(figsize=(10, 5))
    plt.plot(jours, performance_data["perf_marche_ml_min"], color="purple", marker='o', label="Productivité Réelle")  # Trace la productivité réelle
    plt.plot(jours, v_cible["vitesse_cible"], color="red", label="Vitesse Cible")  # Trace la vitesse cible
    plt.title("Productivité de la machine onduleuse en mL/min mois février", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("vitesse", color="blue")
    plt.grid(True, which="both", alpha=0.6)
    plt.xticks(rotation=45, ha='right')
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "productivite.png"))
    plt.show()

    # 6. Graphique comparant la vitesse réelle et la vitesse cible (barres)
    plt.figure(figsize=(10, 5))
    plt.bar(jours_numeric, v_cible["vitesse_cible"], color='green', label="Vitesse Cible", alpha=0.7, width=0.4)  # Barres pour la vitesse cible
    plt.bar(jours_numeric + 0.4, performance_data["perf_marche_ml_min"], color='orange', label="Vitesse Réelle", alpha=0.7, width=0.4)  # Barres pour la vitesse réelle (décalées)
    plt.title('La vitesse reel vs la vitesse cible', fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("vitesse", color="blue")
    plt.xticks(jours_numeric, jours, rotation=45, ha='right')  # Utilise jours_numeric pour positionner les ticks, mais affiche les noms des jours
    plt.legend(loc="lower right")
    plt.grid(axis='y', alpha=0.75)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "vitesse_cible_reelle_barres.png"))
    plt.show()

    # 7. Graphique de l'écart de vitesse
    plt.figure(figsize=(10, 5))
    plt.bar(jours, performance_data["ecart_vitesse"], color="royalblue")  # Barres pour l'écart de vitesse
    plt.title('La valeur d\'ecart entre la vitesse cible et la vitesse reel', fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("Ecart de vitesse", color="blue")
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', alpha=0.75)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "ecart_vitesse.png"))
    plt.show()

    # 8. Graphique des pourcentages d'arrêts techniques et opérationnels (empilées)
    plt.figure(figsize=(10, 5))
    plt.bar(jours, performance_data["% Arrêts Techniques"], color='green', label="Arrêts Techniques")  # Barres pour les arrêts techniques
    plt.bar(jours, performance_data["% Arrêts Opérationnels"], bottom=performance_data["% Arrêts Techniques"], color='orange', label="Arrêts Opérationnels")  # Barres empilées pour les arrêts opérationnels
    plt.title('Pourcentage des arrets technique et operationnels \npar rapport au temps de production planifie', fontsize=12)
    plt.xlabel("jours", color="blue")
    plt.ylabel("% arrêts", color="blue")
    plt.xticks(rotation=45, ha='right')
    plt.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "arrets_techniques_operationnels_empiles.png"))
    plt.show()

    # 9. Graphique des pourcentages d'arrêts techniques (séparé)
    plt.figure(figsize=(10, 5))
    plt.bar(jours, performance_data["% Arrêts Techniques"], color='green', label="Arrêts Techniques")
    plt.title('Pourcentage des arrets techniques \npar rapport au temps de production planifie', fontsize=12)
    plt.xlabel("jours", color="blue")
    plt.ylabel("% arrêts techniques", color="blue")
    plt.xticks(rotation=45, ha='right')
    plt.ylim(0, 0.25)  # Ajuste l'échelle de l'axe Y
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "arrets_techniques.png"))
    plt.show()

    # 10. Graphique des pourcentages d'arrêts opérationnels (séparé)
    plt.figure(figsize=(10, 5))
    plt.bar(jours, performance_data["% Arrêts Opérationnels"], color='orange', label="Arrêts Opérationnels")
    plt.title('Pourcentage des arrets operationnels \npar rapport au temps de production planifie', fontsize=12)
    plt.xlabel("jours", color="blue")
    plt.ylabel("% arrêts opérationnels", color="blue")
    plt.xticks(rotation=45, ha='right')
    plt.ylim(0, 0.25)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "arrets_operationnels.png"))
    plt.show()

    # 11. Graphique du MTBF (Temps Moyen Entre les Pannes)
    plt.figure(figsize=(10, 5))
    plt.plot(jours, performance_data["MTBF"], color="red", marker='o')
    plt.title("L'évolution de MTBF  mois février", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("MTBF", color="blue")
    plt.grid(True, which="both", alpha=0.6)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "MTBF.png"))
    plt.show()

    # 12. Graphique du MTTR (Temps Moyen de Réparation)
    plt.figure(figsize=(10, 5))
    plt.plot(jours, performance_data["MTTR"], color="green", marker='o')
    plt.title("L'évolution de MTTR  mois février", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("MTTR", color="blue")
    plt.grid(True, which="both", alpha=0.6)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "MTTR.png"))
    plt.show()

    # 13. Diagramme de Pareto des causes de perte de productivité
    plt.figure(figsize=(10, 5))
    fig, cs = plt.subplots()  # Crée une figure et un sous-graphique
    plt.xticks(rotation=5)
    plt.grid(True, which="both", alpha=0.6)
    plt.xlabel("causes", color="blue")
    plt.ylabel("pourcentage %", color="blue")
    plt.title("Diagramme de Pareto des causes de perts en term de productivité")
    cs.bar(cause.columns, cause.loc[0], color=["#24d453", "#246dd4", "#d47624"])  # Crée les barres pour les causes
    cs.set_ylim(0, 0.5)  # Ajuste l'échelle de l'axe Y pour les pourcentages
    cum = cs.twinx()  # Crée un deuxième axe Y partageant le même axe X
    cum.plot(cause.columns, [sum(cause.loc[0][:i]) for i in range(1, len(cause.columns) + 1)], color="black", marker="o", label="% cumulé")  # Trace la courbe cumulée
    cum.set_ylim(0, 1.05)  # Ajuste l'échelle de l'axe Y pour le pourcentage cumulé
    cum.set_ylabel("% cumulé", color="blue")
    cum.legend(loc="lower right")
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "pareto.png"))
    plt.show()

    # 14. Graphique comparant la productivité actuelle et améliorée
    plt.figure(figsize=(10, 5))
    plt.plot(jours_numeric, performance_data["perf_marche_ml_min"], marker="o", label="Productivité actuelle")  # Trace la productivité actuelle
    plt.plot(jours_numeric, resultat["amelioration"], marker="x", label="Productivité améliorée")  # Trace la productivité améliorée
    plt.plot(jours_numeric, v_cible["vitesse_cible"], label="La vitesse cible")  # Trace la vitesse cible
    plt.title("Resultats des amelioration à réalisés", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("vitesse(mL/min)", color="blue")
    plt.legend(loc="lower right")
    plt.grid(True, which="both", alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "resultats_amelioration.png"))
    plt.show()

    # 15. Graphique des bénéfices actuels et après améliorations
    plt.figure(figsize=(10, 5))
    plt.bar(jours_numeric, benefices_apres, label="Bénéfices après améliorations", alpha=0.7, width=0.4)  # Barres pour les bénéfices après
    plt.bar(jours_numeric + 0.4, benefices_actuels, label="Bénéfices actuels", alpha=0.7, width=0.4)  # Barres pour les bénéfices actuels (décalées)
    plt.title("Resultats économique des amelioration à réalisés", fontsize=14)
    plt.xlabel("jours", color="blue")
    plt.ylabel("Bénéfices(Millions de Dirhams)", color="blue")
    plt.xticks(jours_numeric, jours, rotation=45, ha='right')  # Utilise jours_numeric pour positionner les ticks, mais affiche les noms des jours
    plt.legend(loc="lower right")
    plt.grid(axis='y', alpha=0.75)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "benefices.png"))
    plt.show()

def analyze_detailed_causes(df3, e_prod_fev):
    """
    Analyse détaillée des causes de perte de productivité à partir de df3.

    Args:
        df3 (pd.DataFrame): DataFrame contenant les données détaillées des causes.
        e_prod_fev (float): La perte de production totale en février.

    Returns:
        tuple: Deux listes - cause_d (dict des causes et leurs pourcentages) et
               commule_d (liste des pourcentages cumulés).
    """

    cause_d = {  # Création d'un dictionnaire pour stocker les causes et leurs contributions
        i: j for i, j in zip(df3["Sub-Cause"], df3["Downtime_hour"] * 210 * 60 / e_prod_fev)
    }
    cause_d["vitesse de la machine"] = 1 - (prod_art_fev + prod_aro_fev) / e_prod_fev  # Ajout de la cause "vitesse de la machine"

    # Tri des causes par pourcentage de contribution (en ordre décroissant)
    cause_d = list(reversed(sorted(cause_d.items(), key=lambda item: item[1])))
    cause_d = {i: j for i, j in cause_d}  # Conversion de la liste triée en dictionnaire

    commule_d = list(map(lambda i: sum(list(cause_d.values())[0:i]), list(range(1, len(cause_d.keys())))))  # Calcul des pourcentages cumulés
    commule_d.append(commule_d[-1])  # Ajout du dernier pourcentage cumulé (100%)

    return cause_d, commule_d


def plot_detailed_pareto(cause_d, commule_d):
    """
    Crée et affiche le diagramme de Pareto détaillé des causes de perte de productivité.

    Args:
        cause_d (dict): Dictionnaire contenant les causes et leurs pourcentages.
        commule_d (list): Liste des pourcentages cumulés.
    """

    plt.figure(figsize=(12, 6))  # Ajuste la taille de la figure pour une meilleure lisibilité
    fig, cs = plt.subplots()  # Crée une figure et un sous-graphique
    plt.xticks(rotation=90, fontsize=8)  # Rotation des étiquettes de l'axe X et réduction de la taille de la police
    plt.grid(True, which="both", alpha=0.6)
    plt.xlabel("causes", color="blue")
    plt.ylabel("pourcentage %", color="blue")
    plt.title("Diagramme de Pareto Détaillé des causes de perts en term de productivité", fontsize=14)

    # Utilisation d'une liste de couleurs aléatoires pour les barres
    cs.bar(list(cause_d.keys())[0:], list(cause_d.values())[0:], color=[get_random_color() for i in range(len(cause_d.keys()))])
    cs.tick_params(axis="x", labelsize=8)  # Ajuste la taille des étiquettes de l'axe X

    cum = cs.twinx()  # Crée un deuxième axe Y partageant le même axe X
    cum.plot(list(cause_d.keys())[0:], commule_d[0:], color="black", marker="o", label="% cumulé")  # Trace la courbe cumulée
    cum.set_ylim(0, 1.05)
    cum.set_ylabel("% cumulé", color="blue")
    cum.legend(loc="lower right")
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "pareto_detaille.png"))
    plt.show()


def load_transformation_data(data_dir):
    """
    Charge les données de performance de la transformation à partir de df4.

    Args:
        data_dir (str): Le chemin vers le répertoire contenant les données.

    Returns:
        tuple: Deux Series pandas - v_cible_t (vitesse cible) et perf_t (performance).
               Retourne None, None en cas d'erreur.
    """
    try:
        df4 = pd.read_excel(os.path.join(os.getcwd(), data_dir, "Performance.xlsx"), sheet_name="Feuil7")
        df4.set_index(df4["Unnamed: 0"], inplace=True)
        df4.columns = range(27)
        v_cible_t = df4.iloc[10]  # Extraction de la vitesse cible
        perf_t = df4.iloc[14]  # Extraction de la performance
        return v_cible_t, perf_t
    except FileNotFoundError:
        print(f"Erreur: Le fichier de données 'Performance.xlsx' n'a pas été trouvé dans '{data_dir}'.")
        return None, None
    except Exception as e:
        print(f"Une erreur s'est produite lors du chargement des données de transformation: {e}")
        return None, None


def plot_onduleuse_transformation_comparison(perf_m2_p, v_cible_m2, perf_t, v_cible_t):
    """
    Compare la productivité de l'onduleuse et de la transformation dans un graphique combiné.

    Args:
        perf_m2_p (Series): Performance de l'onduleuse.
        v_cible_m2 (Series): Vitesse cible de l'onduleuse.
        perf_t (Series): Performance de la transformation.
        v_cible_t (Series): Vitesse cible de la transformation.
    """
    jours_numeric = range(1, len(perf_m2_p) + 1)
    fig, plt1 = plt.subplots(figsize=(12, 6))

    # Productivité de l'onduleuse
    plt1.plot(jours_numeric, perf_m2_p, marker="o", color="blue", label="productivité de l'onduleuse")
    plt1.legend(loc="upper left")
    plt1.set_xlabel("jours", color="blue")
    plt1.set_ylabel("cadence de l'onduleuse (m^2/heure)", color="blue")

    # Productivité de la transformation (axe Y secondaire)
    plt2 = plt1.twinx()
    plt2.plot(jours_numeric, perf_t[1:], marker="x", color="green", label="productivité du transformation")
    plt2.plot(jours_numeric, v_cible_t[1:], color="red", label="productivité cible(m^2/heure)")
    plt2.set_ylabel("cadence du transformation (m^2/heure)", color="green")
    plt2.legend(loc="upper right")

    # Ajustement des limites des axes Y (optionnel, basé sur les données)
    # plt1.set_ylim(min(perf_m2_p), max(perf_m2_p))
    # plt2.set_ylim(min(perf_t[1:]), max(perf_t[1:]))

    plt.title("Comparaison entre la productivité de l'onduleuse et du transformation", fontsize=14)
    plt1.grid(True, which="both", alpha=0.6)
    plt2.grid(True, which="both", alpha=0.6)
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, "comparaison_onduleuse_transformation.png"))
    plt.show()

if __name__ == "__main__":
    # Chargement des données
    df, df2, df3 = load_data(DATA_DIR)
    if df is not None and df2 is not None and df3 is not None:
        # Calcul des indicateurs de performance
        performance_data = calculate_performance_metrics(df)
        if performance_data is not None:
            # Création des DataFrames récapitulatifs
            v_cible, cause, resultat = create_summary_dataframes(performance_data)
            if v_cible is not None and cause is not None and resultat is not None:
                # Calcul du TRS et des bénéfices
                trs_moyen, trs_quotidien = calculate_trs(performance_data, df2.loc[0][2:28]+df2.loc[4][2:28])
                benefices_actuels, benefices_apres = calculate_benefices(performance_data, df2.loc[0][2:28]+df2.loc[4][2:28])

                # Visualisation des données
                plot_performance_graphs(performance_data, v_cible, resultat, trs_quotidien, cause, df2.loc[0][2:28]+df2.loc[4][2:28], benefices_actuels, benefices_apres)

                # Analyse détaillée des causes et diagramme de Pareto
                cause_d, commule_d = analyze_detailed_causes(df3, performance_data["ecart_prod"].sum())
                plot_detailed_pareto(cause_d, commule_d)

                # Comparaison avec les données de transformation (si disponibles)
                v_cible_t, perf_t = load_transformation_data(DATA_DIR)
                if v_cible_t is not None and perf_t is not None:
                    plot_onduleuse_transformation_comparison(performance_data["perf_marche_m2_h"], v_cible_m2, perf_t, v_cible_t)


import streamlit as st
import pandas as pd
import os
from datetime import datetime

# ==========================
# CONFIGURATION PAGE
# ==========================

st.set_page_config(
    page_title="Gestion Restaurant Galette",
    page_icon="🍔",
    layout="wide"
)

# ==========================
# FONCTIONS UTILITAIRES
# ==========================

def charger_excel(fichier, colonnes):

    try:

        if os.path.exists(fichier):

            df = pd.read_excel(fichier)

            # Ajouter les colonnes manquantes
            for col in colonnes:

                if col not in df.columns:
                    df[col] = None

            return df

        else:

            df = pd.DataFrame(
                columns=colonnes
            )

            df.to_excel(
                fichier,
                index=False
            )

            return df

    except:

        df = pd.DataFrame(
            columns=colonnes
        )

        return df


def sauvegarder(df, fichier):

    df.to_excel(
        fichier,
        index=False
    )


# ==========================
# CREATION DES FICHIERS
# ==========================

def creer_fichiers():

    fichiers = {

        "utilisateurs.xlsx":[
            "nom",
            "prenom",
            "identifiant",
            "motdepasse",
            "role"
        ],

        # matières premières
        "matieres.xlsx":[
            "Produit",
            "Quantité",
            "Unité",
            "Seuil"
        ],

        # produits finis
        "produits_finis.xlsx":[
            "Produit",
            "Prix"
        ],

        # portionnement
        "portionnement.xlsx":[
            "Date",
            "Produit",
            "Quantité utilisée",
            "Taille portion",
            "Nombre portions"
        ],

        # recettes
        "recettes.xlsx":[
            "Produit fini",
            "Matière",
            "Quantité"
        ],

        # ventes
        "ventes.xlsx":[
            "Date",
            "Produit",
            "Quantité",
            "Prix",
            "Montant"
        ],

        # mouvements
        "mouvements.xlsx":[
            "Date",
            "Produit",
            "Action",
            "Quantité"
        ],

        # inventaire
        "inventaire.xlsx":[
            "Date",
            "Produit",
            "Stock réel"
        ],
        
        "stock_portions.xlsx":[
            "Produit",
            "Nombre portions"
        ],
        
        # magasin
        "magasin.xlsx":[
            "Date",
            "Produit",
            "Quantité",
            "Unité",
            "Fournisseur"
        ],
        
        "stock_produits_finis.xlsx":[
            "Produit fini",
            "Nombre portions"
        ],

        "fabrication.xlsx":[
            "Date",
            "Produit fini",
            "Quantité fabriquée"
        ],
    }

    for fichier, colonnes in fichiers.items():

        if not os.path.exists(
            fichier
        ):

            pd.DataFrame(
                columns=colonnes
            ).to_excel(
                fichier,
                index=False
            )

    # création admin automatique

    users = pd.read_excel(
        "utilisateurs.xlsx"
    )

    if len(users)==0:

        admin = pd.DataFrame({

            "nom":["Admin"],
            "prenom":["System"],
            "identifiant":["admin"],
            "motdepasse":["admin123"],
            "role":["admin"]

        })

        sauvegarder(
            admin,
            "utilisateurs.xlsx"
        )


creer_fichiers()


# ==========================
# CHARGEMENT DONNEES
# ==========================

users = charger_excel(
    "utilisateurs.xlsx",
    ["nom","prenom","identifiant","motdepasse","role"]
)

matieres = charger_excel(
    "matieres.xlsx",
    ["Produit","Quantité","Unité","Seuil"]
)

produits_finis = charger_excel(
    "produits_finis.xlsx",
    ["Produit","Prix"]
)

portionnement = charger_excel(
    "portionnement.xlsx",
    [
        "Date",
        "Produit",
        "Quantité utilisée",
        "Taille portion",
        "Nombre portions"
    ]
)

recettes = charger_excel(
    "recettes.xlsx",
    [
        "Produit fini",
        "Matière",
        "Quantité"
    ]
)

ventes = charger_excel(
    "ventes.xlsx",
    [
        "Date",
        "Produit",
        "Quantité",
        "Prix",
        "Montant"
    ]
)

mouvements = charger_excel(
    "mouvements.xlsx",
    [
        "Date",
        "Produit",
        "Action",
        "Quantité"
    ]
)

inventaire = charger_excel(
    "inventaire.xlsx",
    [
        "Date",
        "Produit",
        "Stock réel"
    ]
)

stock_portions = charger_excel(
    "stock_portions.xlsx",
    [
        "Produit",
        "Nombre portions"
    ]
)

magasin = charger_excel(
    "magasin.xlsx",
    [
        "Date",
        "Produit",
        "Quantité",
        "Unité",
        "Fournisseur"
    ]
)

stock_produits_finis = charger_excel(
    "stock_produits_finis.xlsx",
    [
        "Produit fini",
        "Nombre portions"
    ]
)

fabrication = charger_excel(
    "fabrication.xlsx",
    [
        "Date",
        "Produit fini",
        "Quantité fabriquée"
    ]
)

# ==========================
# SESSION
# ==========================

if "connecte" not in st.session_state:
    st.session_state.connecte=False

if "utilisateur" not in st.session_state:
    st.session_state.utilisateur=""
    
# ==========================
# NORMALISATION DONNEES
# ==========================

users["identifiant"] = (
    users["identifiant"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.lower()
)

users["motdepasse"] = (
    users["motdepasse"]
    .fillna("")
    .astype(str)
    .str.strip()
)

# ==========================
# SIDEBAR
# ==========================

if not st.session_state.connecte:

    choix = st.sidebar.selectbox(
        "Choisir",
        [
            "Connexion",
            "Inscription"
        ]
    )

else:

    st.sidebar.success(
        f"Connecté : {st.session_state.utilisateur}"
    )

    if st.sidebar.button(
        "Déconnexion"
    ):

        st.session_state.connecte = False
        st.session_state.utilisateur = ""

        st.rerun()


# ==========================
# CONNEXION
# ==========================

if (
    not st.session_state.connecte
    and choix=="Connexion"
):

    st.title(
        "Connexion utilisateur"
    )

    identifiant = st.text_input(
        "Identifiant"
    )

    mdp = st.text_input(
        "Mot de passe",
        type="password"
    )

    if st.button(
        "Se connecter"
    ):

        identifiant = (
            str(identifiant)
            .strip()
            .lower()
        )

        mdp = (
            str(mdp)
            .strip()
        )

        if identifiant=="" or mdp=="":

            st.warning(
                "Remplir tous les champs"
            )

        else:

            utilisateur = users[

                (
                    users[
                        "identifiant"
                    ]
                    ==identifiant
                )

                &

                (
                    users[
                        "motdepasse"
                    ]
                    ==mdp
                )

            ]

            if len(utilisateur)>0:

                st.session_state.connecte=True

                st.session_state.utilisateur=(
                    identifiant
                )

                st.success(
                    "Connexion réussie"
                )

                st.rerun()

            else:

                st.error(
                    "Identifiant ou mot de passe incorrect"
                )


# ==========================
# INSCRIPTION
# ==========================

if (

    not st.session_state.connecte
    and choix=="Inscription"

):

    st.title(
        "Créer un compte"
    )

    nom = st.text_input(
        "Nom"
    )

    prenom = st.text_input(
        "Prénom"
    )

    identifiant = st.text_input(
        "Identifiant"
    )

    mdp = st.text_input(
        "Mot de passe",
        type="password"
    )

    confirmation = st.text_input(
        "Confirmer mot de passe",
        type="password"
    )

    if st.button(
        "Créer compte"
    ):

        nom=nom.strip()
        prenom=prenom.strip()

        identifiant=(
            identifiant
            .strip()
            .lower()
        )

        mdp=mdp.strip()

        # vérification champs

        if (

            nom==""
            or prenom==""
            or identifiant==""
            or mdp==""

        ):

            st.warning(
                "Tous les champs sont obligatoires"
            )

        elif mdp!=confirmation:

            st.error(
                "Les mots de passe sont différents"
            )

        elif identifiant in users[
            "identifiant"
        ].values:

            st.error(
                "Identifiant déjà utilisé"
            )

        else:

            nouveau = pd.DataFrame({

                "nom":[nom],
                "prenom":[prenom],
                "identifiant":[identifiant],
                "motdepasse":[mdp],
                "role":[
                    "utilisateur"
                ]

            })

            users = pd.concat(

                [
                    users,
                    nouveau
                ],

                ignore_index=True

            )

            sauvegarder(

                users,
                "utilisateurs.xlsx"

            )

            st.success(
                "Compte créé avec succès"
            )

            st.rerun()
            
# ==========================
# APPLICATION CONNECTEE
# ==========================

if st.session_state.connecte:

    ligne = users[
        users["identifiant"]
        ==
        st.session_state.utilisateur
    ]

    role = ligne.iloc[0]["role"]

    # ==========================
    # ADMIN
    # ==========================

    if role=="admin":

        st.title(
            "🔐 Administration Restaurant Galette"
        )

        menu = st.sidebar.selectbox(

            "Menu",

            [
                "Magasin",
                "Sortie Magasin",
                "Matières premières",
                "Fabrication Produits Finis",
                "Produits finis",
                "Portionnement",
                "Recettes",
                "Ventes",
                "Consulter Magasin",
                "Consulter matières premières",
                "Consulter Les recettes",
                "Consulter produits",
                "Historique"
            ]
        )

        
        # ==========================
        # Magasin
        # ==========================
        if menu == "Magasin":
    
            st.subheader("Arrivage produits au magasin")

            produit = st.text_input("Produit")

            unite = st.selectbox(
                "Unité",
                ["kg", "g", "L", "ml", "pièce"]
            )

            quantite = st.number_input(
                "Quantité reçue",
                min_value=0.0
            )

            fournisseur = st.text_input(
                "Fournisseur"
            )

            if st.button("Enregistrer arrivage"):

                nv = pd.DataFrame({

                    "Date":[datetime.now()],
                    "Produit":[produit],
                    "Quantité":[quantite],
                    "Unité":[unite],
                    "Fournisseur":[fournisseur]

                })

                magasin = pd.concat(
                    [magasin, nv],
                    ignore_index=True
                )

                sauvegarder(
                    magasin,
                    "magasin.xlsx"
                )

                st.success(
                    "Arrivage enregistré"
                )

                st.rerun()
                
        # ==========================
        # Sortie Magasin
        # ==========================        
        elif menu == "Sortie Magasin":
    
            st.subheader(
                "Transfert vers production"
            )

            produits_disponibles = (
                magasin.groupby("Produit")
                ["Quantité"]
                .sum()
                .reset_index()
            )

            if len(produits_disponibles) == 0:

                st.warning(
                    "Aucun produit en magasin"
                )

            else:

                produit = st.selectbox(
                    "Produit",
                    produits_disponibles["Produit"]
                )

                stock_magasin = produits_disponibles[
                    produits_disponibles["Produit"] == produit
                ].iloc[0]["Quantité"]

                st.info(
                    f"Stock magasin : {stock_magasin}"
                )

                quantite = st.number_input(
                    "Quantité à transférer",
                    min_value=0.0
                )

                if st.button(
                    "Transférer"
                ):

                    if quantite > stock_magasin:

                        st.error(
                            "Stock insuffisant"
                        )

                    else:

                        # Déduire magasin

                        reste = quantite

                        for i in magasin.index:

                            if (
                                magasin.loc[i, "Produit"]
                                == produit
                            ):

                                dispo = magasin.loc[
                                    i,
                                    "Quantité"
                                ]

                                if dispo >= reste:

                                    magasin.loc[
                                        i,
                                        "Quantité"
                                    ] -= reste

                                    break

                                else:

                                    magasin.loc[
                                        i,
                                        "Quantité"
                                    ] = 0

                                    reste -= dispo

                        magasin = magasin[
                            magasin["Quantité"] > 0
                        ]

                        # Ajouter stock matières

                        if produit in matieres["Produit"].values:

                            matieres.loc[
                                matieres["Produit"] == produit,
                                "Quantité"
                            ] += quantite

                        else:

                            unite = magasin[
                                magasin["Produit"] == produit
                            ].iloc[0]["Unité"]

                            nv = pd.DataFrame({

                                "Produit":[produit],
                                "Quantité":[quantite],
                                "Unité":[unite],
                                "Seuil":[0]

                            })

                            matieres = pd.concat(
                                [matieres, nv],
                                ignore_index=True
                            )

                        sauvegarder(
                            magasin,
                            "magasin.xlsx"
                        )

                        sauvegarder(
                            matieres,
                            "matieres.xlsx"
                        )

                        st.success(
                            "Transfert effectué"
                        )

                        st.rerun()
                
        # ==========================
        # AJOUT MATIERE PREMIERE
        # ==========================

        if menu=="Matières premières":

            st.subheader(
                "Ajouter matière première"
                )

            produit = st.text_input(
                "Nom produit"
                )

            unite = st.selectbox(

                "Unité",

                [
                    "kg",
                    "g",
                    "L",
                    "ml",
                    "pièce"
                ]
            )

            quantite = st.number_input(
                "Stock initial",
                min_value=0.0
            )

            seuil = st.number_input(
                "Seuil alerte",
                min_value=0.0
            )

            if st.button(
                "Ajouter matière"
            ):

                produit = produit.strip()

                if produit=="":

                    st.warning(
                        "Nom obligatoire"
                    )

                elif produit.lower() in (

                    matieres[
                        "Produit"
                    ]
                    .astype(str)
                    .str.lower()
                    .values
                    ):

                        st.error(
                            "Produit déjà existant"
                        )

                else:

                    nv = pd.DataFrame({

                        "Produit":[
                            produit
                        ],

                        "Quantité":[
                            quantite
                        ],

                        "Unité":[
                            unite
                            ],

                        "Seuil":[
                            seuil
                            ]

                    })

                    matieres = pd.concat(

                        [
                            matieres,
                            nv
                        ],

                        ignore_index=True
                    )

                    sauvegarder(
                        matieres,
                        "matieres.xlsx"
                    )

                    st.success(
                         "Matière ajoutée"
                    )

                    st.rerun()        
               

        # ==========================
        # AJOUT PRODUIT FINI
        # ==========================

        elif menu=="Produits finis":

            st.subheader(
                "Ajouter produit fini"
            )

            nom = st.text_input(
                "Nom produit"
            )

            prix = st.number_input(
                "Prix vente",
                min_value=0.0
            )

            if st.button(
                "Ajouter produit"
            ):

                nom=nom.strip()

                if nom=="":

                    st.warning(
                        "Nom obligatoire"
                    )

                elif nom.lower() in (

                    produits_finis[
                        "Produit"
                    ]
                    .astype(str)
                    .str.lower()
                    .values

                ):

                    st.error(
                        "Produit déjà existant"
                    )

                else:

                    nv = pd.DataFrame({

                        "Produit":[nom],
                        "Prix":[prix]

                    })

                    produits_finis = pd.concat(

                        [
                            produits_finis,
                            nv
                        ],

                        ignore_index=True
                    )

                    sauvegarder(

                        produits_finis,
                        "produits_finis.xlsx"

                    )

                    st.success(
                        "Produit ajouté"
                    )

                    st.rerun()

        
        # ==========================
        # Fabrication Produits Finis
        # ==========================
        elif menu == "Fabrication Produits Finis":
    
            st.subheader("Fabrication des produits finis")

            produit_fini = st.selectbox(
                "Produit fini",
                produits_finis["Produit"]
            )

            quantite_fabrication = st.number_input(
                "Nombre de portions à fabriquer",
                min_value=1,
                step=1
            )

            recette = recettes[
                recettes["Produit fini"] == produit_fini
            ]

            if len(recette) == 0:

                st.warning(
                    "Aucune recette définie"
                )

            else:

                st.write("Composition de la recette :")
                st.dataframe(recette)

                if st.button("Fabriquer"):

                    stock_ok = True

                    # Vérification des portions disponibles
                    for _, ligne in recette.iterrows():

                        matiere = ligne["Matière"]

                        qte_recette = float(
                            str(ligne["Quantité"]).split()[0]
                        )

                        besoin = (
                            qte_recette *
                            quantite_fabrication
                        )

                        stock_ligne = stock_portions[
                            stock_portions["Produit"] == matiere
                        ]

                        if len(stock_ligne) == 0:

                            st.error(
                                f"{matiere} introuvable"
                            )

                            stock_ok = False
                            break

                        disponible = float(
                            stock_ligne.iloc[0]["Nombre portions"]
                        )

                        if disponible < besoin:

                            st.error(
                                f"Stock insuffisant pour {matiere}"
                            )

                            stock_ok = False
                            break

                    # Déduction des matières
                    if stock_ok:

                        for _, ligne in recette.iterrows():

                            matiere = ligne["Matière"]

                            qte_recette = float(
                                str(ligne["Quantité"]).split()[0]
                            )

                            besoin = (
                                qte_recette *
                                quantite_fabrication
                            )

                            stock_portions.loc[
                                stock_portions["Produit"]
                                == matiere,
                                "Nombre portions"
                            ] -= besoin

                        sauvegarder(
                            stock_portions,
                            "stock_portions.xlsx"
                        )

                        # Ajout produit fini
                        if produit_fini in (
                            stock_produits_finis[
                                "Produit fini"
                            ].values
                        ):

                            stock_produits_finis.loc[
                                stock_produits_finis[
                                    "Produit fini"
                                ] == produit_fini,
                                "Nombre portions"
                            ] += quantite_fabrication

                        else:

                            nv = pd.DataFrame({

                                "Produit fini":[
                                    produit_fini
                                ],

                                "Nombre portions":[
                                    quantite_fabrication
                                ]

                            })

                            stock_produits_finis = pd.concat(
                                [
                                    stock_produits_finis,
                                    nv
                                ],
                                ignore_index=True
                            )

                        sauvegarder(
                            stock_produits_finis,
                            "stock_produits_finis.xlsx"
                        )

                        # Historique fabrication
                        hist = pd.DataFrame({

                            "Date":[datetime.now()],

                            "Produit fini":[
                                produit_fini
                            ],

                            "Quantité fabriquée":[
                                quantite_fabrication
                            ]

                        })

                        fabrication = pd.concat(
                            [fabrication, hist],
                            ignore_index=True
                        )

                        sauvegarder(
                            fabrication,
                            "fabrication.xlsx"
                        )

                        st.success(
                            f"{quantite_fabrication} portions de {produit_fini} fabriquées"
                        )

                        st.rerun()
        
        
        # ==========================
        # PORTIONNEMENT
        # ==========================

        elif menu=="Portionnement":
    
            st.subheader("Transformation en portions (automatique)")

            produit = st.selectbox(
                "Produit",
                matieres["Produit"]
            )

            ligne = matieres[
                matieres["Produit"] == produit
            ]

            stock = ligne.iloc[0]["Quantité"]
            unite = ligne.iloc[0]["Unité"]

            st.info(f"Stock disponible : {stock} {unite}")

            quantite = st.number_input(
                "Quantité à utiliser",
                min_value=0.0
            )

            nb_portions = st.number_input(
                "Nombre de portions souhaitées",
                min_value=1
            )

            if st.button("Créer portions"):

                if quantite > stock:
                    st.error("Stock insuffisant")

                else:

                    # conversion en grammes si besoin
                    if unite == "kg":
                        total_g = quantite * 1000
                    elif unite == "g":
                        total_g = quantite
                    else:
                        total_g = quantite

                    taille_portion = total_g / nb_portions

                    # mise à jour stock matière
                    matieres.loc[
                        matieres["Produit"] == produit,
                        "Quantité"
                    ] -= quantite

                    sauvegarder(matieres, "matieres.xlsx")

                    # historique portionnement
                    nv = pd.DataFrame({
                        "Date": [datetime.now()],
                        "Produit": [produit],
                        "Quantité utilisée": [quantite],
                        "Taille portion": [taille_portion],
                        "Nombre portions": [nb_portions]
                    })

                    portionnement = pd.concat(
                        [portionnement, nv],
                        ignore_index=True
                    )

                    sauvegarder(portionnement, "portionnement.xlsx")

                    # mise à jour stock portions
                    if produit in stock_portions["Produit"].values:

                        stock_portions.loc[
                            stock_portions["Produit"] == produit,
                            "Nombre portions"
                        ] += nb_portions

                    else:

                        stock_portions = pd.concat([
                            stock_portions,
                            pd.DataFrame({
                                "Produit": [produit],
                                "Nombre portions": [nb_portions]
                            })
                        ], ignore_index=True)

                    sauvegarder(stock_portions, "stock_portions.xlsx")

                    st.success(
                        f"Taille d'une portion : {taille_portion:.2f} g | {nb_portions} portions créées"
                    )

                    st.rerun()
        
        
        
        # ==========================
        # VENTES
        # ==========================
        elif menu == "Ventes":
    
            st.title("💰 Gestion des ventes")

            # =========================
            # CHOIX PRODUIT
            # =========================
            produit = st.selectbox(
                "Produit vendu",
                stock_produits_finis["Produit fini"]
            )

            # =========================
            # STOCK DISPONIBLE
            # =========================
            ligne_stock = stock_produits_finis[
                stock_produits_finis["Produit fini"] == produit
            ]

            if len(ligne_stock) == 0:
                st.error("Produit non disponible dans le stock portions")

            else:

                stock = int(ligne_stock.iloc[0]["Nombre portions"])

                st.info(f"📦 Stock disponible : {stock} portions")

                # =========================
                # QUANTITÉ À VENDRE
                # =========================
                quantite = st.number_input(
                    "Nombre de portions à vendre",
                    min_value=1,
                    step=1
                )

                # =========================
                # PRIX UNITAIRE
                # =========================
                prix_ligne = stock_produits_finis[
                    stock_produits_finis["Produit"] == produit
                ]

                if len(prix_ligne) == 0:
                    st.error("Prix non défini pour ce produit")

                else:

                    prix_unitaire = float(prix_ligne.iloc[0]["Prix"])

                    st.success(f"💵 Prix unitaire : {prix_unitaire} FCFA")

                    # =========================
                    # VENTE
                    # =========================
                    if st.button("Valider la vente"):

                        # Vérification stock
                        if quantite > stock:
                            st.error("❌ Stock insuffisant")

                        else:

                            # 1. Mise à jour stock portions
                            stock_portions.loc[
                                stock_portions["Produit"] == produit,
                                "Nombre portions"
                            ] -= quantite

                            sauvegarder(
                                stock_portions,
                                "stock_portions.xlsx"
                            )

                            # 2. Calcul montant
                            montant = quantite * prix_unitaire

                            # 3. Enregistrement vente
                            nouvelle_vente = pd.DataFrame({

                                "Date": [datetime.now()],
                                "Produit": [produit],
                                "Quantité": [quantite],
                                "Prix": [prix_unitaire],
                                "Montant": [montant]

                            })

                            ventes = pd.concat(
                                [ventes, nouvelle_vente],
                                ignore_index=True
                            )

                            sauvegarder(
                                ventes,
                                "ventes.xlsx"
                            )

                            st.success("✔ Vente enregistrée avec succès")

                            st.info(
                                f"💰 Total encaissé : {montant:,.0f} FCFA"
                            )

                            st.rerun()
        
        
        # ==========================
        # AFFICHER MATIERES
        # ==========================

        elif menu=="Consulter matières premières":

            st.subheader(
                "Liste matières premières"
            )

            st.dataframe(
                matieres,
                use_container_width=True
            )
            
            
        # ==========================
        # AFFICHER les recette
        # ==========================

        elif menu=="Consulter Les recettes":

            st.subheader(
                "Liste des recettes"
            )

            st.dataframe(
                recettes,
                use_container_width=True
            )

        # ==========================
        # AFFICHER PRODUITS
        # ==========================

        elif menu=="Consulter produits":

            st.subheader(
                "Liste produits finis"
            )

            st.dataframe(

                produits_finis,
                use_container_width=True

            )


        # ==========================
        # Consulter Magasin
        # ==========================
        elif menu == "Consulter Magasin":
    
            st.subheader(
                "Stock magasin"
            )

            stock = (
                magasin.groupby(
                    ["Produit","Unité"]
                )["Quantité"]
                .sum()
                .reset_index()
            )

            st.dataframe(
                stock,
                use_container_width=True
            )
            
        # ==========================
        # RECETTES
        # ==========================

        elif menu=="Recettes":

            st.subheader(
                "Création recette produit"
            )

            produit_fini = st.selectbox(

                "Produit fini",

                produits_finis["Produit"]

            )

            ingredient = st.selectbox(

                "Matière première",

                matieres["Produit"]

            )

            quantite = st.number_input(

                "Quantité utilisée",

                min_value=0.0

            )

            unite = st.selectbox(

                "Type utilisation",

                [

                    "portion",
                    "g",
                    "kg",
                    "pièce"

                ]

            )

            if st.button(
                "Ajouter recette"
            ):

                recette = pd.DataFrame({

                    "Produit fini":[
                        produit_fini
                    ],

                    "Matière":[
                        ingredient
                    ],

                    "Quantité":[

                        str(
                            quantite
                        )

                        + " "

                        +

                        unite

                    ]

                })

                recettes = pd.concat(

                    [
                        recettes,
                        recette
                    ],

                    ignore_index=True
                )

                sauvegarder(

                    recettes,
                    "recettes.xlsx"

                )

                st.success(
                    "Recette ajoutée"
                )

                st.rerun()

            st.subheader(
                "Recettes existantes"
            )

            filtre = recettes[

                recettes[
                    "Produit fini"
                ]
                ==
                produit_fini

            ]

            st.dataframe(

                filtre,
                use_container_width=True

            )
            
            
        # ==========================
        # HISTORIQUE
        # ==========================

        elif menu == "Historique":
    
            st.title("📊 Historiques du système")

            # =========================
            # CHOIX HISTORIQUE
            # =========================
            choix = st.selectbox(
                "Choisir un historique",
                [
                    "Magasin",
                    "Sorties Magasin",
                    "Portionnement",
                    "Ventes"
                ]
            )

            # =========================
            # FILTRES PAR DATE
            # =========================
            col1, col2 = st.columns(2)

            with col1:
                date_debut = st.date_input(
                    "Date début",
                    value=None
                )

            with col2:
                date_fin = st.date_input(
                    "Date fin",
                    value=None
                )

            # Fonction locale de filtrage
            def filtrer(df, colonne):

                df = df.copy()

                df[colonne] = pd.to_datetime(df[colonne], errors="coerce")

                if date_debut:
                    df = df[df[colonne] >= pd.to_datetime(date_debut)]

                if date_fin:
                    df = df[df[colonne] <= pd.to_datetime(date_fin)]

                return df

            # =========================
            # HISTORIQUE MAGASIN
            # =========================
            if choix == "Magasin":

                st.subheader("📦 Historique des arrivages magasin")

                df = filtrer(magasin, "Date")

                st.dataframe(df, use_container_width=True)

                st.info(f"Total enregistrements : {len(df)}")

            # =========================
            # HISTORIQUE SORTIES MAGASIN
            # =========================
            elif choix == "Sorties Magasin":

                st.subheader("🚚 Historique des sorties magasin")

                sorties = charger_excel(
                    "sorties_magasin.xlsx",
                    [
                        "Date",
                        "Produit magasin",
                        "Matière première",
                        "Quantité",
                        "Utilisateur"
                    ]
                )

                sorties = filtrer(sorties, "Date")

                st.dataframe(sorties, use_container_width=True)

                st.info(f"Total sorties : {len(sorties)}")

            # =========================
            # HISTORIQUE PORTIONNEMENT
            # =========================
            elif choix == "Portionnement":

                st.subheader("🍳 Historique du portionnement")

                df = filtrer(portionnement, "Date")

                st.dataframe(df, use_container_width=True)

                if len(df) > 0:
                    st.success(
                        f"Total portions créées : {df['Nombre portions'].sum()}"
                    )

            # =========================
            # HISTORIQUE VENTES
            # =========================
            elif choix == "Ventes":

                st.subheader("💰 Historique des ventes")

                df = filtrer(ventes, "Date")

                st.dataframe(df, use_container_width=True)

                if len(df) > 0:

                    st.metric(
                        "Chiffre d'affaires total",
                        f"{df['Montant'].sum():,.0f} FCFA"
                    )

                    st.info(
                        f"Total ventes : {len(df)}"
                    )
            
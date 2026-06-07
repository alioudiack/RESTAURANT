import streamlit as st
import pandas as pd
import os
from datetime import datetime
import io

# ==========================
# CONFIGURATION PAGE
# ==========================

st.set_page_config(
    page_title="Restaurant Galette",
    page_icon="🍽️",
    layout="wide"
)

# ==========================
# CSS PERSONNALISÉ
# ==========================

st.markdown("""
<style>
    [data-testid="stSidebar"] {background: #1a1a2e;}
    [data-testid="stSidebar"] * {color: #eee !important;}
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem 1.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 0.5rem;
    }
    .alert-badge {
        background: #ff4757;
        color: white;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 0.75rem;
        font-weight: bold;
    }
    .success-card {
        background: linear-gradient(135deg, #11998e, #38ef7d);
        padding: 1rem;
        border-radius: 10px;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ==========================
# FONCTIONS UTILITAIRES
# ==========================

def charger_excel(fichier, colonnes):
    try:
        if os.path.exists(fichier):
            df = pd.read_excel(fichier)
            for col in colonnes:
                if col not in df.columns:
                    df[col] = None
            return df
        else:
            df = pd.DataFrame(columns=colonnes)
            df.to_excel(fichier, index=False)
            return df
    except:
        return pd.DataFrame(columns=colonnes)


def sauvegarder(df, fichier):
    df.to_excel(fichier, index=False)


def export_csv(df, nom_fichier):
    """Retourne un bouton de téléchargement CSV."""
    csv = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Télécharger CSV",
        data=csv,
        file_name=nom_fichier,
        mime="text/csv"
    )


def compter_alertes(matieres):
    """Compte les matières en dessous du seuil d'alerte."""
    if len(matieres) == 0:
        return 0
    alertes = matieres[
        matieres.apply(
            lambda r: pd.notna(r["Seuil"])
            and pd.notna(r["Quantité"])
            and float(r["Quantité"]) < float(r["Seuil"]),
            axis=1
        )
    ]
    return len(alertes)


# ==========================
# CREATION DES FICHIERS
# ==========================

def creer_fichiers():
    fichiers = {
        "utilisateurs.xlsx": ["nom", "prenom", "identifiant", "motdepasse", "role"],
        "matieres.xlsx": ["Produit", "Quantité", "Unité", "Seuil"],
        "produits_finis.xlsx": ["Produit", "Prix"],
        "portionnement.xlsx": ["Date", "Produit", "Quantité utilisée", "Taille portion", "Nombre portions"],
        "recettes.xlsx": ["Produit fini", "Matière", "Quantité"],
        "ventes.xlsx": ["Date", "Produit", "Quantité", "Prix", "Montant"],
        "mouvements.xlsx": ["Date", "Produit", "Action", "Quantité"],
        "inventaire.xlsx": ["Date", "Produit", "Stock réel"],
        "stock_portions.xlsx": ["Produit", "Nombre portions"],
        "magasin.xlsx": ["Date", "Produit", "Quantité", "Unité", "Fournisseur"],
        "stock_produits_finis.xlsx": ["Produit fini", "Nombre portions"],
        "fabrication.xlsx": ["Date", "Produit fini", "Quantité fabriquée"],
        "sorties_magasin.xlsx": ["Date", "Produit", "Quantité", "Unité"],
    }
    for fichier, colonnes in fichiers.items():
        if not os.path.exists(fichier):
            pd.DataFrame(columns=colonnes).to_excel(fichier, index=False)

    users = pd.read_excel("utilisateurs.xlsx")
    if len(users) == 0:
        admin = pd.DataFrame({
            "nom": ["Admin"], "prenom": ["System"],
            "identifiant": ["admin"], "motdepasse": ["admin123"], "role": ["admin"]
        })
        sauvegarder(admin, "utilisateurs.xlsx")


creer_fichiers()

# ==========================
# CHARGEMENT DONNEES
# ==========================

def charger_tout():
    return {
        "users": charger_excel("utilisateurs.xlsx", ["nom", "prenom", "identifiant", "motdepasse", "role"]),
        "matieres": charger_excel("matieres.xlsx", ["Produit", "Quantité", "Unité", "Seuil"]),
        "produits_finis": charger_excel("produits_finis.xlsx", ["Produit", "Prix"]),
        "portionnement": charger_excel("portionnement.xlsx", ["Date", "Produit", "Quantité utilisée", "Taille portion", "Nombre portions"]),
        "recettes": charger_excel("recettes.xlsx", ["Produit fini", "Matière", "Quantité"]),
        "ventes": charger_excel("ventes.xlsx", ["Date", "Produit", "Quantité", "Prix", "Montant"]),
        "mouvements": charger_excel("mouvements.xlsx", ["Date", "Produit", "Action", "Quantité"]),
        "inventaire": charger_excel("inventaire.xlsx", ["Date", "Produit", "Stock réel"]),
        "stock_portions": charger_excel("stock_portions.xlsx", ["Produit", "Nombre portions"]),
        "magasin": charger_excel("magasin.xlsx", ["Date", "Produit", "Quantité", "Unité", "Fournisseur"]),
        "stock_produits_finis": charger_excel("stock_produits_finis.xlsx", ["Produit fini", "Nombre portions"]),
        "fabrication": charger_excel("fabrication.xlsx", ["Date", "Produit fini", "Quantité fabriquée"]),
        "sorties_magasin": charger_excel("sorties_magasin.xlsx", ["Date", "Produit", "Quantité", "Unité"]),
    }


data = charger_tout()
users = data["users"]
matieres = data["matieres"]
produits_finis = data["produits_finis"]
portionnement = data["portionnement"]
recettes = data["recettes"]
ventes = data["ventes"]
mouvements = data["mouvements"]
inventaire = data["inventaire"]
stock_portions = data["stock_portions"]
magasin = data["magasin"]
stock_produits_finis = data["stock_produits_finis"]
fabrication = data["fabrication"]
sorties_magasin = data["sorties_magasin"]

# ==========================
# NORMALISATION
# ==========================

users["identifiant"] = users["identifiant"].fillna("").astype(str).str.strip().str.lower()
users["motdepasse"] = users["motdepasse"].fillna("").astype(str).str.strip()

# ==========================
# SESSION
# ==========================

if "connecte" not in st.session_state:
    st.session_state.connecte = False
if "utilisateur" not in st.session_state:
    st.session_state.utilisateur = ""

# ==========================
# SIDEBAR
# ==========================

nb_alertes = compter_alertes(matieres)

if not st.session_state.connecte:
    choix = st.sidebar.selectbox("Accès", ["Connexion", "Inscription"])
else:
    st.sidebar.markdown(f"### 👤 {st.session_state.utilisateur}")
    if nb_alertes > 0:
        st.sidebar.markdown(
            f'⚠️ <span class="alert-badge">{nb_alertes} alerte(s) stock</span>',
            unsafe_allow_html=True
        )
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Déconnexion"):
        st.session_state.connecte = False
        st.session_state.utilisateur = ""
        st.rerun()


# ==========================
# CONNEXION
# ==========================

if not st.session_state.connecte and choix == "Connexion":

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("## 🍽️ Restaurant Galette")
        st.markdown("### Connexion")
        identifiant = st.text_input("Identifiant")
        mdp = st.text_input("Mot de passe", type="password")

        if st.button("Se connecter", use_container_width=True):
            identifiant = str(identifiant).strip().lower()
            mdp = str(mdp).strip()
            if identifiant == "" or mdp == "":
                st.warning("Remplir tous les champs")
            else:
                utilisateur = users[
                    (users["identifiant"] == identifiant) &
                    (users["motdepasse"] == mdp)
                ]
                if len(utilisateur) > 0:
                    st.session_state.connecte = True
                    st.session_state.utilisateur = identifiant
                    st.success("Connexion réussie")
                    st.rerun()
                else:
                    st.error("Identifiant ou mot de passe incorrect")


# ==========================
# INSCRIPTION
# ==========================

if not st.session_state.connecte and choix == "Inscription":

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("## Créer un compte")
        nom = st.text_input("Nom")
        prenom = st.text_input("Prénom")
        identifiant = st.text_input("Identifiant")
        mdp = st.text_input("Mot de passe", type="password")
        confirmation = st.text_input("Confirmer mot de passe", type="password")

        if st.button("Créer compte", use_container_width=True):
            nom = nom.strip()
            prenom = prenom.strip()
            identifiant = identifiant.strip().lower()
            mdp = mdp.strip()

            if not all([nom, prenom, identifiant, mdp]):
                st.warning("Tous les champs sont obligatoires")
            elif mdp != confirmation:
                st.error("Les mots de passe ne correspondent pas")
            elif identifiant in users["identifiant"].values:
                st.error("Identifiant déjà utilisé")
            else:
                nouveau = pd.DataFrame({
                    "nom": [nom], "prenom": [prenom],
                    "identifiant": [identifiant], "motdepasse": [mdp],
                    "role": ["utilisateur"]
                })
                users = pd.concat([users, nouveau], ignore_index=True)
                sauvegarder(users, "utilisateurs.xlsx")
                st.success("Compte créé avec succès ! Vous pouvez vous connecter.")


# ==========================
# APPLICATION CONNECTÉE
# ==========================

if st.session_state.connecte:

    ligne_user = users[users["identifiant"] == st.session_state.utilisateur]
    role = ligne_user.iloc[0]["role"]

    # ==========================
    # ADMIN
    # ==========================

    if role == "admin":

        MENU_OPTIONS = [
            "🏠 Tableau de bord",
            "📦 Magasin",
            "🚚 Sortie Magasin",
            "🧂 Matières premières",
            "🏭 Fabrication Produits Finis",
            "🍽️ Produits finis",
            "✂️ Portionnement",
            "📋 Recettes",
            "💰 Ventes",
            "🔍 Consulter Magasin",
            "🔍 Consulter Matières",
            "🔍 Consulter Recettes",
            "🔍 Consulter Produits",
            "📊 Historique",
            "👥 Gestion Utilisateurs",
        ]

        menu = st.sidebar.selectbox("Menu", MENU_OPTIONS)

        # ==========================
        # TABLEAU DE BORD
        # ==========================
        if menu == "🏠 Tableau de bord":

            st.title("🍽️ Tableau de bord — Restaurant Galette")
            st.markdown(f"*Aujourd'hui : {datetime.now().strftime('%d/%m/%Y %H:%M')}*")
            st.markdown("---")

            # KPIs principaux
            col1, col2, col3, col4 = st.columns(4)

            ca_total = ventes["Montant"].sum() if len(ventes) > 0 else 0
            ventes_today = ventes[
                pd.to_datetime(ventes["Date"], errors="coerce").dt.date == datetime.now().date()
            ] if len(ventes) > 0 else pd.DataFrame()
            ca_today = ventes_today["Montant"].sum() if len(ventes_today) > 0 else 0
            total_matieres = len(matieres)
            total_portions = stock_produits_finis["Nombre portions"].sum() if len(stock_produits_finis) > 0 else 0

            with col1:
                st.metric("💰 CA Total", f"{ca_total:,.0f} FCFA")
            with col2:
                st.metric("📅 CA Aujourd'hui", f"{ca_today:,.0f} FCFA")
            with col3:
                st.metric("🧂 Matières premières", total_matieres)
            with col4:
                st.metric("🍽️ Portions disponibles", int(total_portions))

            st.markdown("---")

            col_left, col_right = st.columns(2)

            # Alertes stock
            with col_left:
                st.subheader("⚠️ Alertes stock")
                if len(matieres) > 0:
                    alertes = matieres[
                        matieres.apply(
                            lambda r: pd.notna(r["Seuil"])
                            and pd.notna(r["Quantité"])
                            and float(r["Quantité"]) < float(r["Seuil"]),
                            axis=1
                        )
                    ]
                    if len(alertes) > 0:
                        st.dataframe(
                            alertes[["Produit", "Quantité", "Unité", "Seuil"]],
                            use_container_width=True
                        )
                    else:
                        st.success("✅ Tous les stocks sont au-dessus du seuil")
                else:
                    st.info("Aucune matière première enregistrée")

            # Stock produits finis
            with col_right:
                st.subheader("🍽️ Stock produits finis")
                if len(stock_produits_finis) > 0:
                    st.dataframe(stock_produits_finis, use_container_width=True)
                else:
                    st.info("Aucun produit fini en stock")

            st.markdown("---")

            # Dernières ventes
            st.subheader("🕐 Dernières ventes")
            if len(ventes) > 0:
                dernieres = ventes.tail(10).sort_index(ascending=False)
                st.dataframe(dernieres, use_container_width=True)
            else:
                st.info("Aucune vente enregistrée")


        # ==========================
        # MAGASIN
        # ==========================
        elif menu == "📦 Magasin":
    
            st.subheader("📦 Arrivage produits au magasin")

            type_produit = st.radio("Type de produit", ["Produit existant", "Nouveau produit"])
            
            with st.form("form_magasin"):
                
                if type_produit == "Produit existant":
                    produits_existants = magasin["Produit"].unique()
                    produit = st.selectbox("Sélectionner le produit", produits_existants)
                else:
                    produit = st.text_input("Nom du nouveau produit")

                unite = st.selectbox("Unité", ["kg", "g", "L", "ml", "pièce"])
                quantite = st.number_input("Quantité reçue", min_value=0.0)
                fournisseur = st.text_input("Fournisseur")
                submitted = st.form_submit_button("Enregistrer arrivage")

            if submitted:
                if not produit.strip():
                    st.warning("Le nom du produit est obligatoire")
                elif quantite <= 0:
                    st.warning("La quantité doit être supérieure à 0")
                else:
                    # Enregistrement de l'arrivage dans le magasin
                    nv = pd.DataFrame({
                        "Date": [datetime.now()],
                        "Produit": [produit.strip()],
                        "Quantité": [quantite],
                        "Unité": [unite],
                        "Fournisseur": [fournisseur.strip()]
                    })
                    magasin = pd.concat([magasin, nv], ignore_index=True)
                    sauvegarder(magasin, "magasin.xlsx")
                    
                    # Enregistrement du mouvement
                    nv_mouvement = pd.DataFrame({
                        "Date": [datetime.now()],
                        "Produit": [produit.strip()],
                        "Action": ["Arrivage"],
                        "Quantité": [quantite]
                    })
                    mouvements = pd.concat([mouvements, nv_mouvement], ignore_index=True)
                    sauvegarder(mouvements, "mouvements.xlsx")
                    
                    st.success(f"✅ Arrivage de {quantite} {unite} de {produit} enregistré")
                    st.rerun()

        
        # ==========================
        # SORTIE MAGASIN
        # ==========================
        elif menu == "🚚 Sortie Magasin":

            st.subheader("🚚 Transfert magasin → production")

            produits_disponibles = magasin.groupby(["Produit", "Unité"])["Quantité"].sum().reset_index()
            produits_disponibles = produits_disponibles[produits_disponibles["Quantité"] > 0]

            if len(produits_disponibles) == 0:
                st.warning("Aucun produit disponible en magasin")
            else:
                produit = st.selectbox("Produit au magasin", produits_disponibles["Produit"])

                ligne_prod = produits_disponibles[produits_disponibles["Produit"] == produit].iloc[0]
                stock_magasin = ligne_prod["Quantité"]
                unite_prod = ligne_prod["Unité"]

                st.info(f"Stock magasin : **{stock_magasin} {unite_prod}**")

                matiere_associee = st.selectbox("Associer à la matière première", matieres["Produit"])

                quantite = st.number_input("Quantité à transférer", min_value=0.0, max_value=float(stock_magasin))

                if st.button("Transférer vers production"):
                    if quantite <= 0:
                        st.warning("Quantité invalide")
                    elif quantite > stock_magasin:
                        st.error("Stock insuffisant")
                    else:
                        reste = quantite
                        for i in magasin.index:
                            if magasin.loc[i, "Produit"] == produit and reste > 0:
                                dispo = magasin.loc[i, "Quantité"]
                                if dispo >= reste:
                                    magasin.loc[i, "Quantité"] -= reste
                                    reste = 0
                                else:
                                    magasin.loc[i, "Quantité"] = 0
                                    reste -= dispo

                        sauvegarder(magasin, "magasin.xlsx")

                        matieres.loc[matieres["Produit"] == matiere_associee, "Quantité"] += quantite
                        sauvegarder(matieres, "matieres.xlsx")

                        sortie = pd.DataFrame({
                            "Date": [datetime.now()],
                            "Produit Magasin": [produit],
                            "Matière Première Cible": [matiere_associee],
                            "Quantité": [quantite],
                            "Unité": [unite_prod]
                        })
                        sorties_magasin = pd.concat([sorties_magasin, sortie], ignore_index=True)
                        sauvegarder(sorties_magasin, "sorties_magasin.xlsx")

                        # Enregistrement du mouvement de sortie
                        mvt_sortie = pd.DataFrame({
                            "Date": [datetime.now()],
                            "Produit": [produit],
                            "Action": ["Sortie"],
                            "Quantité": [quantite]
                        })
                        mouvements = pd.concat([mouvements, mvt_sortie], ignore_index=True)
                        sauvegarder(mouvements, "mouvements.xlsx")

                        st.success(f"✅ {quantite} {unite_prod} de {produit} transférés vers {matiere_associee}")
                        st.rerun()


        # ==========================
        # MATIÈRES PREMIÈRES
        # ==========================
        elif menu == "🧂 Matières premières":

            st.subheader("🧂 Ajouter une matière première")

            with st.form("form_matiere"):
                produit = st.text_input("Nom du produit")
                col1, col2 = st.columns(2)
                with col1:
                    unite = st.selectbox("Unité", ["kg", "g", "L", "ml", "pièce"])
                    quantite = st.number_input("Stock initial", min_value=0.0)
                with col2:
                    seuil = st.number_input("Seuil d'alerte", min_value=0.0)
                submitted = st.form_submit_button("Ajouter matière")

            if submitted:
                produit = produit.strip()
                if not produit:
                    st.warning("Nom obligatoire")
                elif produit.lower() in matieres["Produit"].astype(str).str.lower().values:
                    st.error("Produit déjà existant")
                else:
                    nv = pd.DataFrame({
                        "Produit": [produit], "Quantité": [quantite],
                        "Unité": [unite], "Seuil": [seuil]
                    })
                    matieres = pd.concat([matieres, nv], ignore_index=True)
                    sauvegarder(matieres, "matieres.xlsx")
                    st.success("✅ Matière ajoutée")
                    st.rerun()

            st.markdown("---")
            st.subheader("Stock actuel")
            if len(matieres) > 0:
                # Colorier les lignes en alerte
                def highlight_alerte(row):
                    if pd.notna(row["Seuil"]) and pd.notna(row["Quantité"]):
                        if float(row["Quantité"]) < float(row["Seuil"]):
                            return ["background-color: #ffe0e0"] * len(row)
                    return [""] * len(row)
                st.dataframe(
                    matieres.style.apply(highlight_alerte, axis=1),
                    use_container_width=True
                )
            else:
                st.info("Aucune matière première enregistrée")


        # ==========================
        # PRODUITS FINIS
        # ==========================
        elif menu == "🍽️ Produits finis":

            st.subheader("🍽️ Ajouter un produit fini")

            with st.form("form_produit"):
                nom = st.text_input("Nom du produit")
                prix = st.number_input("Prix de vente (FCFA)", min_value=0.0)
                submitted = st.form_submit_button("Ajouter produit")

            if submitted:
                nom = nom.strip()
                if not nom:
                    st.warning("Nom obligatoire")
                elif nom.lower() in produits_finis["Produit"].astype(str).str.lower().values:
                    st.error("Produit déjà existant")
                else:
                    nv = pd.DataFrame({"Produit": [nom], "Prix": [prix]})
                    produits_finis = pd.concat([produits_finis, nv], ignore_index=True)
                    sauvegarder(produits_finis, "produits_finis.xlsx")
                    st.success("✅ Produit ajouté")
                    st.rerun()

            st.markdown("---")
            st.subheader("Liste des produits")
            if len(produits_finis) > 0:
                # Modifier les prix
                produit_edit = st.selectbox("Modifier le prix d'un produit", ["— Sélectionner —"] + list(produits_finis["Produit"]))
                if produit_edit != "— Sélectionner —":
                    prix_actuel = float(produits_finis[produits_finis["Produit"] == produit_edit].iloc[0]["Prix"])
                    nouveau_prix = st.number_input("Nouveau prix (FCFA)", value=prix_actuel, min_value=0.0)
                    if st.button("Mettre à jour le prix"):
                        produits_finis.loc[produits_finis["Produit"] == produit_edit, "Prix"] = nouveau_prix
                        sauvegarder(produits_finis, "produits_finis.xlsx")
                        st.success("Prix mis à jour")
                        st.rerun()
                st.dataframe(produits_finis, use_container_width=True)
            else:
                st.info("Aucun produit fini enregistré")


        # ==========================
        # FABRICATION
        # ==========================
        elif menu == "🏭 Fabrication Produits Finis":
            st.subheader("🏭 Fabrication des produits finis")
            if len(produits_finis) == 0:
                st.warning("Aucun produit fini défini")
            else:
                produit_fini = st.selectbox("Produit fini à fabriquer", produits_finis["Produit"])
                quantite_fabrication = st.number_input("Nombre de portions à fabriquer", min_value=1, step=1)
                
                # Récupération de la recette du produit
                recette = recettes[recettes["Produit fini"] == produit_fini]

                if len(recette) == 0:
                    st.warning("⚠️ Aucune recette définie pour ce produit")
                else:
                    st.markdown("**Composition et Vérification des stocks de matières premières :**")
                    faisable = True
                    
                    # On vérifie chaque ingrédient
                    for _, lig in recette.iterrows():
                        matiere = lig["Matière"]
                        # Analyse de la quantité requise dans la recette
                        quantite_recette_str = str(lig["Quantité"])
                        parties = quantite_recette_str.split()[0]
                        qte_recette = float(parties) if parties.replace('.', '', 1).isdigit() else 0
                        besoin = qte_recette * quantite_fabrication
                        
                        # Vérification dans la table des portions
                        stock_ligne = stock_portions[stock_portions["Produit"] == matiere]
                        disponible = float(stock_ligne.iloc[0]["Nombre portions"]) if len(stock_ligne) > 0 else 0.0
                        
                        statut = "✅" if disponible >= besoin else "❌"
                        if disponible < besoin:
                            faisable = False
                        st.write(f"{statut} **{matiere}** — Besoin : {besoin} | Disponible : {disponible}")

                    if st.button("Fabriquer", disabled=not faisable):
                        # Déduction du stock de portions
                        for _, lig in recette.iterrows():
                            matiere = lig["Matière"]
                            parties = str(lig["Quantité"]).split()[0]
                            qte_recette = float(parties) if parties.replace('.', '', 1).isdigit() else 0
                            besoin = qte_recette * quantite_fabrication
                            
                            stock_portions.loc[
                                stock_portions["Produit"] == matiere,
                                "Nombre portions"
                            ] -= besoin
                        
                        sauvegarder(stock_portions, "stock_portions.xlsx")

                        # Ajout du produit fini dans le stock
                        if produit_fini in stock_produits_finis["Produit fini"].values:
                            stock_produits_finis.loc[
                                stock_produits_finis["Produit fini"] == produit_fini,
                                "Nombre portions"
                            ] += quantite_fabrication
                        else:
                            nv = pd.DataFrame({
                                "Produit fini": [produit_fini],
                                "Nombre portions": [quantite_fabrication]
                            })
                            stock_produits_finis = pd.concat([stock_produits_finis, nv], ignore_index=True)
                        
                        sauvegarder(stock_produits_finis, "stock_produits_finis.xlsx")
                        
                        # Ajout à l'historique
                        hist = pd.DataFrame({
                            "Date": [datetime.now()],
                            "Produit fini": [produit_fini],
                            "Quantité fabriquée": [quantite_fabrication]
                        })
                        fabrication = pd.concat([fabrication, hist], ignore_index=True)
                        sauvegarder(fabrication, "fabrication.xlsx")
                        
                        st.success(f"✅ {quantite_fabrication} portions de **{produit_fini}** fabriquées")
                        st.rerun()

                    if not faisable:
                        st.error("❌ Stock insuffisant pour lancer la fabrication")

        # ==========================
        # PORTIONNEMENT
        # ==========================
        elif menu == "✂️ Portionnement":
            st.subheader("✂️ Création de portions")

            if len(matieres) == 0:
                st.warning("Aucune matière première disponible")
            else:
                produit = st.selectbox("Produit à portionner", matieres["Produit"])
                lig = matieres[matieres["Produit"] == produit].iloc[0]
                stock = float(lig["Quantité"])
                unite = lig["Unité"]

                st.metric("Stock disponible en cuisine", f"{stock} {unite}")
                
                # Saisie du nombre de portions et de la quantité utilisée
                nb_portions = st.number_input("Nombre de portions souhaitées", min_value=1, step=1)
                quantite = st.number_input(
                    f"Quantité brute à utiliser ({unite})", 
                    min_value=0.0, 
                    max_value=float(stock)
                )

                if quantite > 0 and nb_portions > 0:
                    taille = (quantite * 1000 if unite == "kg" else quantite) / nb_portions
                    st.info(f"Taille d'une portion : **{taille:.2f} {'g' if unite in ['kg','g'] else unite}**")

                if st.button("Créer les portions"):
                    if quantite > stock:
                        st.error("Stock insuffisant pour cette opération")
                    elif quantite <= 0:
                        st.warning("Quantité invalide")
                    else:
                        total = (quantite * 1000 if unite == "kg" else quantite)
                        taille_portion = total / nb_portions

                        # 1. Déduction du stock de matières premières
                        matieres.loc[matieres["Produit"] == produit, "Quantité"] -= quantite
                        sauvegarder(matieres, "matieres.xlsx")

                        # 2. Enregistrement dans l'historique de portionnement
                        nv = pd.DataFrame({
                            "Date": [datetime.now()], 
                            "Produit": [produit],
                            "Quantité utilisée": [quantite], 
                            "Taille portion": [taille_portion],
                            "Nombre portions": [nb_portions]
                        })
                        portionnement = pd.concat([portionnement, nv], ignore_index=True)
                        sauvegarder(portionnement, "portionnement.xlsx")

                        # 3. Ajout des portions dans le stock dédié
                        if produit in stock_portions["Produit"].values:
                            stock_portions.loc[stock_portions["Produit"] == produit, "Nombre portions"] += nb_portions
                        else:
                            stock_portions = pd.concat([
                                stock_portions,
                                pd.DataFrame({"Produit": [produit], "Nombre portions": [nb_portions]})
                            ], ignore_index=True)

                        sauvegarder(stock_portions, "stock_portions.xlsx")
                        
                        st.success(f"✅ {nb_portions} portions créées avec succès !")
                        st.rerun()


        # ==========================
        # RECETTES
        # ==========================
        elif menu == "📋 Recettes":

            st.subheader("📋 Gestion des recettes")

            if len(produits_finis) == 0:
                st.warning("Aucun produit fini défini")
            elif len(matieres) == 0:
                st.warning("Aucune matière première définie")
            else:
                with st.form("form_recette"):
                    produit_fini = st.selectbox("Produit fini", produits_finis["Produit"])
                    ingredient = st.selectbox("Matière première", matieres["Produit"])
                    col1, col2 = st.columns(2)
                    with col1:
                        quantite = st.number_input("Quantité", min_value=0.0)
                    with col2:
                        unite = st.selectbox("Unité", ["portion", "g", "kg", "pièce"])
                    submitted = st.form_submit_button("Ajouter à la recette")

                if submitted:
                    recette_nv = pd.DataFrame({
                        "Produit fini": [produit_fini],
                        "Matière": [ingredient],
                        "Quantité": [f"{quantite} {unite}"]
                    })
                    recettes = pd.concat([recettes, recette_nv], ignore_index=True)
                    sauvegarder(recettes, "recettes.xlsx")
                    st.success("✅ Recette ajoutée")
                    st.rerun()

                st.markdown("---")
                produit_filtre = st.selectbox("Voir la recette de :", ["Tous"] + list(produits_finis["Produit"]))
                filtre = recettes if produit_filtre == "Tous" else recettes[recettes["Produit fini"] == produit_filtre]
                st.dataframe(filtre, use_container_width=True)

                # Suppression d'un ingrédient
                if len(recettes) > 0:
                    st.markdown("---")
                    st.subheader("Supprimer un ingrédient de recette")
                    idx = st.selectbox("Sélectionner la ligne à supprimer", recettes.index,
                                       format_func=lambda i: f"{recettes.loc[i,'Produit fini']} — {recettes.loc[i,'Matière']} ({recettes.loc[i,'Quantité']})")
                    if st.button("Supprimer cette ligne"):
                        recettes = recettes.drop(index=idx).reset_index(drop=True)
                        sauvegarder(recettes, "recettes.xlsx")
                        st.success("Ligne supprimée")
                        st.rerun()


        # ==========================
        # VENTES
        # ==========================
        elif menu == "💰 Ventes":

            st.title("💰 Enregistrer une vente")

            if len(stock_produits_finis) == 0:
                st.warning("Aucun produit fini en stock")
            else:
                produit = st.selectbox("Produit vendu", stock_produits_finis["Produit fini"])

                ligne_stock = stock_produits_finis[stock_produits_finis["Produit fini"] == produit]
                stock = int(ligne_stock.iloc[0]["Nombre portions"]) if len(ligne_stock) > 0 else 0

                # Récupérer le prix dans produits_finis (pas stock_produits_finis)
                ligne_prix = produits_finis[produits_finis["Produit"] == produit]
                prix_unitaire = float(ligne_prix.iloc[0]["Prix"]) if len(ligne_prix) > 0 else 0.0

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("📦 Stock disponible", f"{stock} portions")
                with col2:
                    st.metric("💵 Prix unitaire", f"{prix_unitaire:,.0f} FCFA")

                quantite = st.number_input("Nombre de portions à vendre", min_value=1, step=1, max_value=max(1, stock))

                if quantite > 0:
                    st.info(f"💰 Total : **{quantite * prix_unitaire:,.0f} FCFA**")

                if st.button("✅ Valider la vente"):
                    if quantite > stock:
                        st.error("❌ Stock insuffisant")
                    elif prix_unitaire == 0:
                        st.warning("⚠️ Prix non défini pour ce produit")
                    else:
                        montant = quantite * prix_unitaire

                        stock_produits_finis.loc[
                            stock_produits_finis["Produit fini"] == produit,
                            "Nombre portions"
                        ] -= quantite
                        sauvegarder(stock_produits_finis, "stock_produits_finis.xlsx")

                        nouvelle_vente = pd.DataFrame({
                            "Date": [datetime.now()],
                            "Produit": [produit],
                            "Quantité": [quantite],
                            "Prix": [prix_unitaire],
                            "Montant": [montant]
                        })
                        ventes = pd.concat([ventes, nouvelle_vente], ignore_index=True)
                        sauvegarder(ventes, "ventes.xlsx")

                        st.success(f"✅ Vente enregistrée — **{montant:,.0f} FCFA** encaissés")
                        st.balloons()
                        st.rerun()


        # ==========================
        # CONSULTER MAGASIN
        # ==========================
        elif menu == "🔍 Consulter Magasin":

            st.subheader("📦 Stock magasin")
            if len(magasin) > 0:
                stock = magasin.groupby(["Produit", "Unité"])["Quantité"].sum().reset_index()
                stock = stock[stock["Quantité"] > 0]
                st.dataframe(stock, use_container_width=True)
                export_csv(stock, "stock_magasin.csv")
            else:
                st.info("Magasin vide")


        # ==========================
        # CONSULTER MATIÈRES
        # ==========================
        elif menu == "🔍 Consulter Matières":

            st.subheader("🧂 Matières premières")
            if len(matieres) > 0:
                st.dataframe(matieres, use_container_width=True)
                export_csv(matieres, "matieres.csv")
            else:
                st.info("Aucune matière première")


        # ==========================
        # CONSULTER RECETTES
        # ==========================
        elif menu == "🔍 Consulter Recettes":

            st.subheader("📋 Recettes")
            if len(recettes) > 0:
                st.dataframe(recettes, use_container_width=True)
                export_csv(recettes, "recettes.csv")
            else:
                st.info("Aucune recette définie")


        # ==========================
        # CONSULTER PRODUITS
        # ==========================
        elif menu == "🔍 Consulter Produits":

            st.subheader("🍽️ Produits finis")

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Catalogue produits**")
                st.dataframe(produits_finis, use_container_width=True)
            with col2:
                st.markdown("**Stock portions disponibles**")
                st.dataframe(stock_produits_finis, use_container_width=True)

            export_csv(produits_finis, "produits_finis.csv")


        # ==========================
        # HISTORIQUE
        # ==========================
        elif menu == "📊 Historique":

            st.title("📊 Historiques")

            choix_hist = st.selectbox("Choisir un historique", [
                "Magasin (arrivages)",
                "Sorties Magasin",
                "Portionnement",
                "Fabrication",
                "Ventes"
            ])

            col1, col2 = st.columns(2)
            with col1:
                date_debut = st.date_input("Date début", value=None)
            with col2:
                date_fin = st.date_input("Date fin", value=None)

            def filtrer(df, colonne):
                df = df.copy()
                df[colonne] = pd.to_datetime(df[colonne], errors="coerce")
                if date_debut:
                    df = df[df[colonne].dt.date >= date_debut]
                if date_fin:
                    df = df[df[colonne].dt.date <= date_fin]
                return df

            if choix_hist == "Magasin (arrivages)":
                st.subheader("📦 Arrivages magasin")
                df = filtrer(magasin, "Date")
                st.dataframe(df, use_container_width=True)
                st.info(f"Total : {len(df)} enregistrement(s)")
                export_csv(df, "historique_magasin.csv")

            elif choix_hist == "Sorties Magasin":
                st.subheader("🚚 Sorties magasin")
                df = filtrer(sorties_magasin, "Date")
                st.dataframe(df, use_container_width=True)
                st.info(f"Total : {len(df)} sortie(s)")
                export_csv(df, "historique_sorties.csv")

            elif choix_hist == "Portionnement":
                st.subheader("✂️ Portionnement")
                df = filtrer(portionnement, "Date")
                st.dataframe(df, use_container_width=True)
                if len(df) > 0:
                    st.success(f"Total portions créées : {df['Nombre portions'].sum()}")
                export_csv(df, "historique_portionnement.csv")

            elif choix_hist == "Fabrication":
                st.subheader("🏭 Fabrication")
                df = filtrer(fabrication, "Date")
                st.dataframe(df, use_container_width=True)
                if len(df) > 0:
                    st.success(f"Total fabriqué : {df['Quantité fabriquée'].sum()} portions")
                export_csv(df, "historique_fabrication.csv")

            elif choix_hist == "Ventes":
                st.subheader("💰 Ventes")
                df = filtrer(ventes, "Date")
                st.dataframe(df, use_container_width=True)
                if len(df) > 0:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Chiffre d'affaires", f"{df['Montant'].sum():,.0f} FCFA")
                    with col2:
                        st.metric("Nombre de ventes", len(df))
                export_csv(df, "historique_ventes.csv")


        # ==========================
        # GESTION UTILISATEURS
        # ==========================
        elif menu == "👥 Gestion Utilisateurs":

            st.title("👥 Gestion des utilisateurs")

            st.dataframe(
                users[["nom", "prenom", "identifiant", "role"]],
                use_container_width=True
            )

            st.markdown("---")
            st.subheader("Modifier le rôle")
            users_non_admin = users[users["identifiant"] != "admin"]
            if len(users_non_admin) > 0:
                user_select = st.selectbox(
                    "Utilisateur",
                    users_non_admin["identifiant"],
                    format_func=lambda x: f"{users[users['identifiant']==x].iloc[0]['prenom']} {users[users['identifiant']==x].iloc[0]['nom']} ({x})"
                )
                nouveau_role = st.selectbox("Nouveau rôle", ["utilisateur", "admin"])
                if st.button("Changer le rôle"):
                    users.loc[users["identifiant"] == user_select, "role"] = nouveau_role
                    sauvegarder(users, "utilisateurs.xlsx")
                    st.success("Rôle mis à jour")
                    st.rerun()

            st.markdown("---")
            st.subheader("Supprimer un utilisateur")
            if len(users_non_admin) > 0:
                user_sup = st.selectbox(
                    "Utilisateur à supprimer",
                    users_non_admin["identifiant"],
                    key="sup",
                    format_func=lambda x: f"{users[users['identifiant']==x].iloc[0]['prenom']} {users[users['identifiant']==x].iloc[0]['nom']} ({x})"
                )
                if st.button("🗑️ Supprimer", type="primary"):
                    users = users[users["identifiant"] != user_sup]
                    sauvegarder(users, "utilisateurs.xlsx")
                    st.success("Utilisateur supprimé")
                    st.rerun()
            else:
                st.info("Aucun utilisateur à gérer")


    # ==========================
    # UTILISATEUR STANDARD
    # ==========================

    else:

        st.title("🍽️ Restaurant Galette")
        st.markdown(f"Bienvenue, **{st.session_state.utilisateur}** !")

        menu_user = st.sidebar.selectbox(
            "Menu",
            [
                "💰 Ventes",
                "🔍 Consulter Stock",
            ]
        )

        if menu_user == "💰 Ventes":

            st.subheader("💰 Enregistrer une vente")

            if len(stock_produits_finis) == 0:
                st.warning("Aucun produit disponible")
            else:
                produit = st.selectbox("Produit", stock_produits_finis["Produit fini"])
                ligne_stock = stock_produits_finis[stock_produits_finis["Produit fini"] == produit]
                stock = int(ligne_stock.iloc[0]["Nombre portions"]) if len(ligne_stock) > 0 else 0
                ligne_prix = produits_finis[produits_finis["Produit"] == produit]
                prix_unitaire = float(ligne_prix.iloc[0]["Prix"]) if len(ligne_prix) > 0 else 0.0

                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Stock", f"{stock} portions")
                with col2:
                    st.metric("Prix unitaire", f"{prix_unitaire:,.0f} FCFA")

                quantite = st.number_input("Quantité", min_value=1, step=1, max_value=max(1, stock))

                if quantite > 0:
                    st.info(f"Total : **{quantite * prix_unitaire:,.0f} FCFA**")

                if st.button("✅ Valider"):
                    if quantite > stock:
                        st.error("Stock insuffisant")
                    else:
                        montant = quantite * prix_unitaire
                        stock_produits_finis.loc[
                            stock_produits_finis["Produit fini"] == produit,
                            "Nombre portions"
                        ] -= quantite
                        sauvegarder(stock_produits_finis, "stock_produits_finis.xlsx")

                        nouvelle_vente = pd.DataFrame({
                            "Date": [datetime.now()], "Produit": [produit],
                            "Quantité": [quantite], "Prix": [prix_unitaire], "Montant": [montant]
                        })
                        ventes = pd.concat([ventes, nouvelle_vente], ignore_index=True)
                        sauvegarder(ventes, "ventes.xlsx")
                        st.success(f"✅ Vente enregistrée — {montant:,.0f} FCFA")
                        st.balloons()
                        st.rerun()

        elif menu_user == "🔍 Consulter Stock":
            st.subheader("🔍 Produits disponibles")
            if len(stock_produits_finis) > 0:
                st.dataframe(stock_produits_finis, use_container_width=True)
            else:
                st.info("Aucun produit disponible")
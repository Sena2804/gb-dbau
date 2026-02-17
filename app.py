import streamlit as st

# C'est aussi simple que ça pour afficher du texte
st.title("Ma première App Streamlit 🚀")
st.write("Si tu vois ça, c'est que ça fonctionne !")

# Un petit widget pour tester l'interactivité
nom = st.text_input("Quel est ton nom ?")
if nom:
    st.success(f"Bienvenue dans l'aventure, {nom} !")
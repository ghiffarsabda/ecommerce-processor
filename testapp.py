import streamlit as st

st.title("Test App")
st.write("If you can see this, the deployment is working!")

# Add a simple button
if st.button("Click me"):
    st.write("Button clicked!")

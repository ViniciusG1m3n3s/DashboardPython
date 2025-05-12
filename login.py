import streamlit as st


# Dicionário com usuários e senhas
usuarios = {"usuario@itau": "senha1", "usuario2@bv": "senha2", "sidnei@itau": "f1nch", "carol@itau": "f1nch", "controller@itau": "f1nch", "viviane@bv": "f1nch", "raissa@maestro": "f1nch", "sidnei@oficios": "f1nch", "viviane_bv@bv": "finch", "bianca@amil": "f1nch", "vinicius@amil": "f1nch", "andrew@unimed": "f1nch"} 

# Função para autenticar usuário
def autenticar(usuario, senha):
    return usuarios.get(usuario) == senha

# Função para carregar o dashboard dependendo do domínio do usuário
def login():
    st.logo("finch.png")
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")

    if st.sidebar.button("Entrar", icon=":material/login:", use_container_width=True, type="primary"):
        if autenticar(usuario, senha):
            st.session_state.logado = True
            st.session_state.usuario_logado = usuario    # Armazena o usuário logado
            st.sidebar.success("Login bem-sucedido!")
            return True  # Retorna True se o login for bem-sucedido
        else:
            st.sidebar.error("Usuário ou senha incorretos.")

    background_image_css = """
    <style>
    [data-testid="stAppViewContainer"] {
        background: 
            linear-gradient(
                to left,
                rgba(211, 33, 0, 0.7),
                rgba(175, 15, 26, 0.7),
                rgba(136, 12, 33, 0.7),
                rgba(96, 15, 33, 0.7),
                rgba(54, 4, 38, 0.7),
                rgba(84, 0, 117, 0.7)
            ),
            url("https://i.pinimg.com/originals/ef/c3/85/efc3857811ea0813316475b79af16c7c.jpg");
        background-size: cover;
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
    }
        @keyframes gradientAnimation {
        0% {
            background-position: 0% 50%;
        }
        50% {
            background-position: 100% 50%;
        }
        100% {
            background-position: 0% 50%;
        }
    }
    header {
    background-color: rgba(255, 255, 255, 0); /* Torna o fundo do cabeçalho transparente */
    color: transparent; /* Remove o texto do cabeçalho (opcional) */
    box-shadow: none; /* Remove a sombra (opcional) */
    }
    [data-testid="stHeader"] {
        background-color: rgba(255, 255, 255, 0); /* Transparente no novo identificador */
    }
    
        .stToolbarActions st-emotion-cache-1p1m4ay e3i9eg820 {
        background-color: rgba(255, 255, 255, 0); /* Transparente no novo identificador */
    }
    
    .stAppToolbar st-emotion-cache-15ecox0 e10jh26i2 {
        background-color: rgba(255, 255, 255, 0); /* Transparente no novo identificador */
    }
    
        /* Seleciona a barra do Streamlit */
    .st-emotion-cache-15ecox0 {
        background-color: rgba(255, 255, 255, 0.0) !important; /* Transparente */
        box-shadow: none !important; /* Remove sombra */
    }

    /* Remove bordas e sombras adicionais */
    .stToolbarActions, .st-emotion-cache-czk5ss {
        background-color: rgba(255, 255, 255, 0.0) !important;
        box-shadow: none !important;
    }
    
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK {
        display: none;
    }
    
    #MainMenu {
        visibility: hidden;
        }   
    #GithubIcon {
    visibility: hidden;
    }
    </style>
    """
    st.markdown(background_image_css, unsafe_allow_html=True)

    hide_streamlit_style = """
                    <style>
                    div[data-testid="stToolbar"] {
                    visibility: hidden;
                    height: 0%;
                    position: fixed;
                    }
                    div[data-testid="stDecoration"] {
                    visibility: hidden;
                    height: 0%;
                    position: fixed;
                    }
                    div[data-testid="stStatusWidget"] {
                    visibility: hidden;
                    height: 0%;
                    position: fixed;
                    }
                    #MainMenu {
                    visibility: hidden;
                    height: 0%;
                    }
                    header {
                    visibility: hidden;
                    height: 0%;
                    }
                    footer {
                    visibility: hidden;
                    height: 0%;
                    }
                    </style>
                    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)
    
    return False  # Retorna False se o login falhar



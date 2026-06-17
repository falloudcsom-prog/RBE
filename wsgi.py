from app import app, initialiser_base

# Garantir l'initialisation même si le module-level a échoué
initialiser_base()

if __name__ == "__main__":
    app.run()

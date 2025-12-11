from flask import Flask, jsonify
from flasgger import Swagger
from .routes import bp as main_bp
from .db_utilis import client, db  # import both client and db
from flask_cors import CORS

def create_app():
    app = Flask(__name__)
    swagger = Swagger(app)
    app.config.from_mapping(SECRET_KEY="dev")
    CORS(app, resources={r"/api/*": {"origins": "*"}})
    
    app.register_blueprint(main_bp)
    
    @app.route("/")
    def home():
        return "Hello, World!"

    @app.get("/db")
    def db_now():
        """Test MongoDB connection"""
        try:
            server_status = db.command("serverStatus")
            return jsonify({
                "status": "connected",
                "localTime": server_status["localTime"],
                "databases": client.list_database_names()  # ✅ call on client, not db
            })
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    
    @app.get("/health")
    def health():
        return jsonify({
            "status": "ok",
            "service": "FarmXpert API"
        })
    
    return app

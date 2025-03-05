from app import app

if __name__ == "__main__":
    # Ensure debug mode is enabled and the server is accessible
    app.run(host='0.0.0.0', port=5000, debug=True)
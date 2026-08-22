from main_app import app, application

if __name__ == '__main__':
    try:
        from waitress import serve
        print("==================================================")
        print(" Server Ujian SMK Budi Murni 2 Berjalan!")
        print(" Akses Lokal  : http://localhost:5000")
        print("==================================================")
        serve(app, host='0.0.0.0', port=5000, threads=32)
    except Exception as e:
        app.run(host='0.0.0.0', port=5000, debug=True)

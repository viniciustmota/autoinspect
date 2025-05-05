from flask import Flask, render_template, request
import os
import main  # Importe o main.py, onde você tem a função de automação

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/executar', methods=['POST'])
def executar():
    matricula = request.form['matricula']
    arquivo = request.files['arquivo']

    # Caminho temporário onde o arquivo será salvo
    caminho_arquivo = os.path.join('uploads', arquivo.filename)
    os.makedirs('uploads', exist_ok=True)
    arquivo.save(caminho_arquivo)

    # Rodar automação com os dados fornecidos
    main.rodar_automacao(matricula, caminho_arquivo)
    return f"Automação iniciada para a matrícula: {matricula}"

if __name__ == '__main__':
     # O Gunicorn cuidará de rodar o servidor, portanto não é necessário deixar o app.run()
    pass

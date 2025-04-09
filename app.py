from flask import Flask, render_template, request, jsonify
import random

app = Flask(__name__)

# Conhecimento sobre Excel
excel_knowledge = {
    "introducao": [
        "Excel é um programa de planilha eletrônica da Microsoft, parte do pacote Office",
        "O Excel serve para organizar dados, fazer cálculos, criar gráficos e análises",
        "É usado para finanças, contabilidade, engenharia e qualquer área que trabalhe com dados",
        "Planilhas Excel usam linhas, colunas e células (como A1, B2) para organizar informações"
    ],
    "atalhos": [
        "Ctrl+N - Novo arquivo | Ctrl+S - Salvar | Ctrl+P - Imprimir",
        "Ctrl+Z - Desfazer | Ctrl+Y - Refazer | Ctrl+C/V/X - Copiar/Colar/Cortar",
        "F2 - Editar célula | F4 - Repetir última ação ou travar referência ($A$1)",
        "Ctrl+Seta - Navegação rápida entre células com dados",
        "Ctrl+Shift+L - Ativar/desativar filtros | Alt+= - AutoSoma",
        "Ctrl+1 - Formatação de células | Ctrl+Shift+$ - Formatar como moeda",
        "Ctrl+Shift+# - Formatar como data | Ctrl+Shift+@ - Formatar como hora",
        "Ctrl+PageUp/PageDown - Alternar entre planilhas | Alt+F1 - Criar gráfico rápido"
    ],
    "funcoes": [
        "=SOMA(A1:A10) - Soma valores em um intervalo",
        "=MÉDIA(B1:B10) - Calcula a média aritmética",
        "=SE(A1>10,\"Aprovado\",\"Reprovado\") - Teste condicional",
        "=PROCV(valor,intervalo,coluna,[falso]) - Busca vertical",
        "=CONT.SE(intervalo,critério) - Conta células que atendem a condição",
        "=SOMASE(intervalo,critério,intervalo_soma) - Soma com condição",
        "=TEXTO(valor,\"formato\") - Formata números como texto",
        "=HOJE() - Data atual | =AGORA() - Data e hora atuais"
    ],
    "dicas": [
        "Use Fórmulas Nomeadas para facilitar referências complexas",
        "Pressione Alt+Enter para quebrar linha dentro de uma célula",
        "Use CONCATENAR ou & para unir textos de diferentes células",
        "Formatar como Tabela (Ctrl+T) facilita a organização dos dados",
        "Experimente Gráficos Dinâmicos para análise visual de dados"
    ]
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/chat', methods=['POST'])
def chat():
    user_input = request.form['user_input'].lower().strip(" ?!")
    
    # Mapeamento inteligente de palavras-chave
    if any(word in user_input for word in ["excel", "planilha", "o que é", "serve", "utiliz", "para que"]):
        response = random.choice(excel_knowledge["introducao"])
    elif any(word in user_input for word in ["atalho", "tecla", "comando", "rapido", "rápido", "ctrl", "shift", "f2", "f4"]):
        response = random.choice(excel_knowledge["atalhos"])
    elif any(word in user_input for word in ["funç", "formula", "fórmula", "calcular", "soma", "media", "média", "proc", "se(", "cont.se", "hoje()"]):
        response = random.choice(excel_knowledge["funcoes"])
    elif any(word in user_input for word in ["dica", "truque", "sugestão", "melhorar"]):
        response = random.choice(excel_knowledge["dicas"])
    elif any(word in user_input for word in ["obrigad", "valeu", "agradeço"]):
        response = "De nada! Estou aqui para ajudar com suas dúvidas de Excel."
    elif any(word in user_input for word in ["oi", "olá", "bom dia", "boa tarde", "boa noite"]):
        response = "Olá! Sou seu assistente de Excel. Posso te explicar sobre:\n- Conceitos básicos\n- Atalhos\n- Funções\n- Dicas profissionais\nO que você quer saber?"
    else:
        if any(word in user_input for word in ["dado", "analis", "gráfico", "pivot"]):
            response = "Para análise de dados, recomendo:\n1. Tabelas Dinâmicas\n2. Funções como =SOMASE()\n3. Gráficos de Dispersão\nQuer detalhes sobre algum?"
        else:
            response = "Posso te ajudar com:\n1. Conceitos básicos do Excel\n2. Atalhos úteis\n3. Funções importantes\n4. Dicas profissionais\nSobre qual assunto quer saber?"
    
    return jsonify({"response": response})

if __name__ == '__main__':
    app.run(debug=True)
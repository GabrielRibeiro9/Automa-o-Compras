# 🤖 Robô de Follow-up de Compras


Este projeto é basicamente um robô que tenta sobreviver ao apocalipse de follow-ups manuais de compras. Para fazer isso, ele assume o controle, identifica os pedidos críticos e alerta quem precisa ser alertado!

Enquanto o robô está "vivo" e rodando, ele gera relatórios em PDF e envia alertas por WhatsApp para garantir que nenhum prazo seja perdido.

> Cuidado com os pedidos muito atrasados! Eles são os "chefões" do jogo.

---

### Por quê?

Este projeto nasceu da necessidade de aplicar meus conhecimentos em Python e automação (RPA) para resolver um problema real e chato: o acompanhamento manual de pedidos. O objetivo era criar uma solução simples e eficaz para economizar tempo, reduzir erros e me introduzir no mundo da automação de processos, tentando respeitar as boas práticas de programação!

Sinta-se à vontade para usar este projeto como quiser, seja para estudo ou uso comercial. Apenas tenha atenção aos dados sensíveis e credenciais!

### Como ele funciona na prática?

O robô segue uma sequência de sobrevivência bem definida:

1.  **Garantir a Comunicação (Selenium):** Primeiro, ele verifica se a nossa principal arma, o serviço de mensagens UltraMsg, está online.
2.  **Mapear o Terreno (Pandas):** O robô faz uma cópia segura da planilha de compras e a analisa para identificar dois tipos de "ameaças": Pedidos Vencendo e Pedidos Atrasados.
3.  **Alertar os Aliados (API WhatsApp):** Para cada fornecedor com pedidos prestes a vencer, o robô envia uma mensagem amigável pelo WhatsApp.
4.  **Criar o Relatório de Batalha (FPDF2):** Com a lista de todos os pedidos "zumbis" (atrasados), o robô gera um relatório de status em PDF.
5.  **Reportar à Base (SMTPLib):** Por fim, ele envia o relatório em PDF por e-mail para a gestão.

### O que eu preciso para rodar?

Você só precisa de um "PC batata" e seguir estes passos:

1.  **Clone este repositório.**
2.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Configure as credenciais** no arquivo `.env` e no código.
4.  **Execute o script:**
    ```bash
    python app.py
    ```

### Tem ideias ou sugestões?

Me mande um e-mail: **[gabrielrdsouza9@gmail.com](mailto:gabrielrdsouza9@gmail.com)**

---

### Considerações Finais

> Dedicado às noites que não dormi porque estava viciado em codificar essa automação e resolver os bugs.
>
> Tenha um ótimo dia!

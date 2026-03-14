🚀 Motor de Precificação Top-Down (Mercado Livre)
Um algoritmo de precificação dinâmica e integração de ERP construído sob medida para o ecossistema do Mercado Livre. Desenvolvido para lidar com a alta complexidade do sistema tributário brasileiro (ICMS-ST, DIFAL, Tese do Século e IPI) cruzado com a matriz tridimensional de custos logísticos do marketplace.

🎯 O Problema Resolvido
No e-commerce brasileiro de autopeças e kits complexos, o preço de venda dita a sobrevivência da operação. Modelos tradicionais de precificação ("Por Dentro" ou Mark-up simples) falham catastroficamente ao não preverem:

    - A variação abrupta de custos logísticos de acordo com o peso, preço (Barreira dos R$ 79) e reputação da conta.
    - A bitributação de impostos em produtos com Substituição Tributária (ICMS-ST).
    - A exclusão do ICMS da base do PIS/COFINS (Tese do Século).
    - A deflação de impostos "Por Fora" (IPI) para integração correta das tags XML (vProd, vIPI) em ERPs.

Este motor inverte a lógica: em vez de somar custos e rezar pela margem, ele recebe a Margem Líquida Alvo e roda um "Loop de Tiers" para encontrar o menor preço de venda possível que garanta 100% do lucro estipulado.

🧠 Arquitetura e Core Features
A arquitetura foi dividida em três camadas distintas executadas via Google Apps Script (V8):

    1. Liquidificador Fiscal (Camada de Dados)
    Quando um "Kit" (Composto) é detectado, o motor aglutina os componentes, seus pesos, custos e naturezas tributárias (Origem, IPI). Ele gera um Bloco Virtual com taxas efetivas e alíquotas ponderadas, permitindo que produtos de origens fiscais distintas (Ex: Nacional Isento + Importado ST) sejam vendidos juntos sem furo de margem.

    2. Motor Financeiro e DRE Dinâmica
    O núcleo matemático do algoritmo. Cruza os dados do Bloco Virtual com o Regime Tributário do seller (Simples Nacional, Lucro Presumido ou Lucro Real com estorno de PIS/COFINS).

        - Tratamento de IPI: Conversão algébrica da alíquota "Por Fora" para uma Taxa Efetiva "Por Dentro", expurgando o imposto da Receita Bruta.
        - Auditoria em Tempo Real: Retorna uma matriz de 12 colunas com a DRE completa da venda (Custo, Comissão, Frete, ICMS, DIFAL, FECOP, PIS/COFINS, IPI, IRPJ, CSLL e Margem Líquida).

    3. Staging Table para ERP (Rateio Proporcional)
    Para resolver a cardinalidade de "1 SKU de Anúncio para N SKUs de Componentes", o motor gera uma tabela de integração normalizada (TGF_VUNCOM). Ele explode o anúncio e rateia financeiramente o Preço Final e o Frete entre as peças, calculando os reais exatos e devolvendo as tags prontas para o XML da NF-e (cProd, qCom, vUnCom, vProd, vIPI, vFrete).

⚙️ Estrutura do Banco de Dados (In-Memory)
    - CONFIG_SELLER: Definição de regime tributário, alíquotas federais e reputação.
    - TGFPRO: Catálogo master de produtos, dimensões, custos e IPI.
    - TGFKIT: A receita de bolo para formação de kits/combos.
    - TGFADS: Painel de controle do pricing, definição de táticas (Frete Forçado) e DRE visual.
    - TGF_VUNCOM: Tabela de saída normalizada para consumo via API (Sankhya, Bling, Tiny).

💻 Instalação e Deploy
O projeto é gerenciado localmente e integrado ao Google Sheets via clasp (Command Line Apps Script Projects).

    1. Clone o repositório.
    2. Faça o login na sua conta Google: clasp login
    3. Configure o arquivo .clasp.json com o Script ID da sua planilha alvo.
    4. Suba o código: clasp push

Desenvolvido com precisão matemática para operações de alta performance pela 360 Gestão.
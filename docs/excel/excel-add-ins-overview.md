# <a name="excel-add-ins-overview"></a>Visão geral dos suplementos do Excel

Um suplemento do Excel permite estender a funcionalidade do aplicativo Excel em várias plataformas, incluindo Office para Windows, Office Online, Office para Mac e Office para iPad. Use suplementos do Excel em uma pasta de trabalho para:

- Interagir com objetos do Excel, ler e gravar dados do Excel. 
- Estender a funcionalidade usando o painel de tarefas ou o painel conteúdo baseado na Web 
- Adicionar botões personalizados da faixa de opções ou itens de menu contextuais
- Fornecer interação mais rica usando janela de caixa de diálogo 

A plataforma Suplementos do Office fornece a estrutura e as APIs JavaScript Office.js que permitem criar e executar suplementos do Excel. Usando a plataforma Suplementos do Office para criar o suplemento do Excel, você receberá os seguintes benefícios:

* **Suporte à plataforma cruzada**: Os suplementos do Excel são executados no Office para Windows, Mac, iOS e Office Online.
* **Implantação centralizada**: Os administradores podem implantar rápida e facilmente suplementos do Excel para usuários em toda uma organização.
* **SSO (logon único)**: Integre facilmente seu suplemento do Excel ao Microsoft Graph.
* **Uso da tecnologia da Web padrão**: Crie um suplemento do Excel usando tecnologias da Web conhecidas, como HTML, CSS e JavaScript.
* **Distribuição pela Office Store**: Compartilhe o suplemento do Excel com uma ampla audiência publicando-o na [Office Store](https://store.office.com/en-us/appshome.aspx).

> **Observação**: Os suplementos são diferentes dos suplementos de COM e VSTO, que são anteriores às soluções de integração do Office que são executadas apenas no Office para Windows. Diferentemente dos suplementos de COM, os suplementos do Excel não exigem a instalação de código no dispositivo de um usuário, nem no Excel. 

## <a name="components-of-an-excel-add-in"></a>Componentes de um suplemento do Excel 

Um suplemento do Excel inclui dois componentes básicos: um aplicativo Web e um arquivo de configuração, chamado de arquivo de manifesto. 

O aplicativo Web usa a [API JavaScript para Office](../../reference/javascript-api-for-office.md) para interagir com objetos no Excel e também pode facilitar a interação com recursos online. Por exemplo, um suplemento pode executar alguma das seguintes tarefas:

* Criar, ler, atualizar e excluir dados na pasta de trabalho (planilhas, intervalos, tabelas, gráficos, itens nomeados e muito mais).
* Executar autorização de usuário em um serviço online usando o fluxo padrão OAuth 2.0.
* Emitir solicitações de API ao Microsoft Graph ou qualquer outra API.

O aplicativo Web pode ser hospedado em qualquer servidor Web, além de poder ser criado usando estruturas do lado do cliente (como Angular, React, jQuery) ou tecnologias do lado do servidor (como ASP.NET, Node.js, PHP).

O [manifesto](../overview/add-in-manifests.md) é um arquivo de configuração XML que define como o suplemento integra-se aos clientes do Office, especificando configurações e recursos, como: 

* A URL do aplicativo Web do suplemento.
* O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.
* Como o suplemento integra-se ao Excel, incluindo qualquer interface de usuário personalizada que o suplemento cria (botões da faixa de opções, menus de contexto, etc.).
* Permissões exigidas pelo suplemento, como leitura e gravação no documento.

Para permitir que os usuários finais instalem e usem um suplemento do Excel, você deve publicar o respectivo manifesto na Office Store ou em um catálogo de Suplementos. 

## <a name="capabilities-of-an-excel-add-in"></a>Recursos de um suplemento do Excel

Além de interagir com o conteúdo da pasta de trabalho, os suplementos do Excel podem adicionar botões personalizados da faixa de opções ou comandos de menu, inserir painéis de tarefas, abrir caixas de diálogo e, até mesmo, inserir objetos sofisticados baseados na web, como gráficos ou visualizações interativas, em uma planilha, conforme mostrado nas capturas de tela a seguir. Para saber mais sobre cada um desses recursos, confira [Estender funcionalidade do Excel](excel-add-ins-extend-excel.md).

**Botões personalizados da faixa de opções**

![Comandos de suplemento](../images/Excel_add-in_commands_Script-Lab.png)

**Painel de tarefas**

![Painel de tarefas do suplemento](../images/Excel_add-in_task_pane_Insights.png)

**Caixa de diálogo**

![Caixa de diálogo do suplemento](../images/Excel_add-in_dialog_choose-number.png)

**Suplemento de conteúdo**

![Suplemento de conteúdo](../images/Excel_add-in_content_map.png)

## <a name="javascript-apis-to-interact-with-workbook-content"></a>APIs JavaScript para interagir com o conteúdo da pasta de trabalho

Um suplemento do Excel interage com objetos no Excel usando a [API JavaScript para Office](../../reference/javascript-api-for-office.md), que inclui dois modelos de objeto JavaScript:

* **API JavaScript do Excel**: Introduzida com o Office 2016, a [API JavaScript do Excel](../../reference/excel/excel-add-ins-reference-overview.md) fornece objetos do Excel fortemente tipados que você pode usar para acessar planilhas, intervalos, tabelas, gráficos e muito mais. 

* **API compartilhada**: Introduzida com o Office 2013, a API compartilhada permite acessar recursos como interface de usuário, caixas de diálogo e configurações de cliente, que são comuns entre vários tipos de aplicativos host, como Word, Excel e PowerPoint. Como a API compartilhada fornece funcionalidade limitada para interação do Excel, você poderá usá-la se seu suplemento precisar ser executado no Excel 2013.

## <a name="next-steps"></a>Próximas etapas

Introdução à [criação de seu primeiro suplemento do Excel](excel-add-ins-get-started-overview.md). Em seguida, saiba mais sobre os [principais conceitos](excel-add-ins-core-concepts.md) da criação de suplementos do Excel.

## <a name="additional-resources"></a>Recursos adicionais

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Práticas recomendadas para o desenvolvimento de Suplementos do Office](../overview/add-in-development-best-practices.md)
- [Diretrizes de design para suplementos do Office](../design/add-in-design.md)
- [Principais conceitos da API JavaScript do Excel](excel-add-ins-core-concepts.md)
- [Referência da API JavaScript do Excel](../../reference/excel/excel-add-ins-reference-overview.md)

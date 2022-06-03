---
title: Amostras de código de suplemento do Office
description: Uma lista de exemplos de código de suplementos do Office para ajudá-lo a aprender e criar seus próprios suplementos.
ms.date: 06/02/2022
localization_priority: high
ms.openlocfilehash: 1ad5fcb3ed860832b093dc6aef212e9d1176f298
ms.sourcegitcommit: 5e678f87b6b886949cc0fcec73468a41fa39fd06
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/03/2022
ms.locfileid: "65872012"
---
# <a name="office-add-in-code-samples"></a>Amostras de código de suplemento do Office

Esses exemplos de código são escritos para ajudá-lo a aprender como usar vários recursos ao desenvolver suplementos do Office.

## <a name="getting-started"></a>Introdução

Os exemplos a seguir mostram como construir o Suplemento do Office mais simples com apenas um manifesto, página da web HTML e um logotipo. Esses componentes são as partes fundamentais de um Suplemento do Office. Para obter informações adicionais sobre os primeiros passos, consulte nossos [primeiros passos](../quickstarts/excel-quickstart-jquery.md) e [tutoriais](/search/?terms=tutorial&scope=Office%20Add-ins).

- [Suplemento "Hello World" do Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/excel-hello-world)
- [Suplemento "Hello world" do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)
- [Suplemento "Olá, mundo" do PowerPoint](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-hello-world)
- [Suplemento do Word "Olá, mundo"](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/word-hello-world)

<br>

---

---

## <a name="excel"></a>Excel

| Name                | Descrição         |
|:--------------------|:--------------------|
| [Abrir no Teams](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-open-in-teams) | Crie uma nova planilha do Excel no Microsoft Teams contendo os dados que você definir.|
| [Inserir um arquivo Excel externo e preenchê-lo com dados JSON](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-insert-file)  | Insira um modelo existente de um arquivo externo do Excel na pasta de trabalho do Excel aberta no momento. Em seguida, preencha o modelo com dados de um serviço Web JSON. |
| [Crie guias contextuais personalizadas na faixa de opções](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/office-contextual-tabs) | Crie uma guia contextual personalizada na faixa de opções na interface do usuário do Office. O exemplo cria uma tabela e, quando o usuário move o foco dentro da tabela, a guia personalizada é exibida. Quando o usuário sai da tabela, a guia personalizada fica oculta. |
| [Use os atalhos do teclado para ações do Suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-keyboard-shortcuts) | Configure um projeto de suplemento básico do Excel que utiliza atalhos de teclado. |
| [Exemplo de função personalizada usando web worker](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/web-worker) | Use web workers em funções personalizadas para evitar o bloqueio da interface do usuário do suplemento do Office. |
| [Use técnicas de armazenamento para acessar dados de um suplemento do Office quando estiver offline](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/Excel.OfflineStorageAddin) | Implemente o localStorage para habilitar a funcionalidade limitada do Suplemento do Office quando um usuário perder a conexão. |
| [Padrão de lote de função personalizada](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/Batching)| Agrupe várias chamadas em uma única chamada para reduzir o número de chamadas de rede para um serviço remoto.|

## <a name="outlook"></a>Outlook

| Nome                | Descrição         |
|:--------------------|:--------------------|
| [Criptografar anexos, processar os participantes de solicitação de reunião e reaja a alterações de data/hora do compromisso](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/outlook-encrypt-attachments) | Use a ativação baseada em eventos para criptografar anexos quando adicionados pelo usuário. Use também a manipulação de eventos para destinatários alterados em uma solicitação de reunião e alterações na data ou hora de início ou de término em uma solicitação de reunião. |
| [Use a ativação baseada em eventos do Outlook para marcar destinatários externos](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external) | Use a ativação baseada em eventos para executar um suplemento do Outlook quando o usuário alterar os destinatários ao redigir uma mensagem. O suplemento também usa a API `appendOnSendAsync` para adicionar um aviso de isenção. |
| [Use a ativação baseada em eventos do Outlook para definir a assinatura](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature) | Use a ativação baseada em eventos para executar um suplemento do Outlook quando o usuário criar uma nova mensagem ou compromisso. O suplemento pode responder a eventos, mesmo quando o painel de tarefas não está aberto. Ele também usa a API `setSignatureAsync`. |
| [Usar Alertas Inteligentes do Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories) | Use os Alertas Inteligentes do Outlook para verificar se as categorias de cores necessárias são aplicadas a uma nova mensagem ou compromisso antes de enviá-la. |

## <a name="word"></a>Word

| Name                | Descrição         |
|:--------------------|:--------------------|
| [Obter, editar e definir conteúdo OOXML em um documento do Word com um suplemento do Word](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-get-set-edit-openxml) | Este exemplo mostra como obter, editar e definir conteúdo OOXML em um documento do Word. O complemento de exemplo fornece um bloco de rascunho para obter o Office Open XML para seu próprio conteúdo e testar seus próprios trechos de código editados do Office Open XML.|
| [Carregar e gravar Open XML no seu suplemento do Word](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/word-add-in-load-and-write-open-xml)  | Este exemplo de suplemento mostra como adicionar uma variedade de tipos de conteúdos avançados a um documento do Word usando o método setSelectedDataAsync com tipo de coerção ooxml. O suplemento também oferece a capacidade de mostrar a marcação do Office Open XML para cada tipo de conteúdo de exemplo na página. |

<br>

---

---

## <a name="authentication-authorization-and-single-sign-on-sso"></a>Autenticação, autorização e logon único (SSO)

| Nome                | Descrição         |
|:--------------------|:--------------------|
| [Suplemento de amostra do Outlook de logon único (SSO)](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) | Use o recurso SSO do Office para fornecer ao suplemento acesso aos dados do Microsoft Graph.|
| [Obtenha dados do OneDrive usando Microsoft Graph e msal.js em um suplemento do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React) | Crie um suplemento do Office, como um aplicativo de página única (SPA) sem back-end, que se conecta ao Microsoft Graph e acesse pastas de trabalho armazenadas no OneDrive for Business para atualizar uma planilha.  |
| [Autenticação do suplemento do Office para o Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET) | Aprenda a criar um suplemento do Microsoft Office que se conecte ao Microsoft Graph e acesse pastas de trabalho armazenadas no OneDrive for Business para atualizar uma planilha. |
| [Autenticação do suplemento do Outlook para Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET). | Crie um suplemento do Outlook que se conecte ao Microsoft Graph e acesse pastas de trabalho armazenadas no OneDrive for Business para redigir uma nova mensagem de email. |
| [Suplemento do Office de Logon único (SSO) com ASP.NET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO) | Use o API `getAccessToken` no Office.js para dar ao suplemento acesso a dados do Microsoft Graph. Este exemplo é criado com base no ASP.NET. |
| [Suplemento Office dee Logon único (SSO) com Node.js](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO) | Use o API `getAccessToken` no Office.js para dar ao suplemento acesso a dados do Microsoft Graph. Este exemplo é criado no Node.js.|

## <a name="shared-javascript-runtime"></a>Tempo de execução de JavaScript compartilhado

| Nome                | Descrição         |
|:--------------------|:--------------------|
| [Compartilhe dados globais com um tempo de execução compartilhado](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-global-state) | Configure um projeto básico que usa o tempo de execução compartilhado para executar código para botões da faixa de opções, painel de tarefas e funções personalizadas em um único tempo de execução do navegador. |
| [Gerencie a faixa de opções e a interface do usuário do painel de tarefas e execute o código no documento aberto](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario) | Crie os botões contextuais da faixa de opções que são ativados com base no estado do seu suplemento. |

<br>

---

---

## <a name="additional-samples"></a>Amostras adicionais

| Nome                | Descrição         |
|:--------------------|:--------------------|
| [Use uma biblioteca compartilhada para migrar seu suplemento do Visual Studio Tools para Office para um suplemento da web do Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/VSTO-shared-code-migration) | Fornece uma estratégia para reutilização de código ao migrar de suplementos do VSTO para suplementos do Office. |
| [Integre uma função do Azure à sua função personalizada do Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AzureFunction) | Integre funções do Azure com funções personalizadas para mover para a nuvem ou integrar serviços adicionais. |
| [Amostras de código DPI dinâmico](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/dynamic-dpi) | Uma coleção de amostras para lidar com alterações de DPI em suplementos COM, VSTO e Office. |

## <a name="next-steps"></a>Próximos passos

Participe do Programa do Desenvolvedor do Microsoft 365. Obtenha uma área restrita, ferramentas e outros recursos gratuitos que você precisa para criar soluções para a plataforma Microsoft 365.

- [A área restrita de desenvolvedor grátis](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Obtenha uma assinatura de desenvolvedor Microsoft 365 E5 gratuita e renovável por 90 dias.
- [Amostra de pacotes de dados](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Configure automaticamente sua área restrita instalando dados de usuário e conteúdo para ajudá-lo a construir suas soluções.
- [Acesso a especialistas](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Acesse eventos da comunidade para aprender com especialistas do Microsoft 365.
- [Recomendações personalizadas](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Localize os recursos para desenvolvedores rapidamente em seu painel personalizado.

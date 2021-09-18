---
title: Amostras de código de suplemento do Office
description: Uma lista de exemplos de código de suplementos do Office para ajudá-lo a aprender e criar seus próprios suplementos.
ms.date: 09/09/2021
localization_priority: high
ms.openlocfilehash: fb595273fa890c6eb16dbfe03fe102a2a3ee6a9a
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/18/2021
ms.locfileid: "59443796"
---
# <a name="office-add-in-code-samples"></a>Amostras de código de suplemento do Office

Esses exemplos de código são escritos para ajudá-lo a aprender como usar vários recursos ao desenvolver suplementos do Office.

## <a name="outlook"></a>Outlook

| Nome                | Descrição         |
|:--------------------|:--------------------|
| [Use a ativação baseada em eventos do Outlook para marcar destinatários externos (visualização)](/samples/officedev/pnp-officeaddins/outlook-add-in-tag-external-recipients) | Use a ativação baseada em eventos para executar um suplemento do Outlook quando o usuário alterar os destinatários ao redigir uma mensagem. O suplemento também usa a API `appendOnSendAsync` para adicionar um aviso de isenção. |
| [Use a ativação baseada em eventos do Outlook para definir a assinatura](/samples/officedev/pnp-officeaddins/outlook-add-in-set-signature/) | Use a ativação baseada em eventos para executar um suplemento do Outlook quando o usuário criar uma nova mensagem ou compromisso. O suplemento pode responder a eventos, mesmo quando o painel de tarefas não está aberto. Ele também usa a API `setSignatureAsync`. |

## <a name="excel"></a>Excel

| Name                | Descrição         |
|:--------------------|:--------------------|
| [Abrir no Teams](/samples/officedev/pnp-officeaddins/office-excel-add-in-open-in-teams/) | Crie uma nova planilha do Excel no Microsoft Teams contendo os dados que você definir.|
| [Crie guias contextuais personalizadas na faixa de opções](/samples/officedev/pnp-officeaddins/office-add-in-contextual-tabs/) | Crie uma guia contextual personalizada na faixa de opções na interface do usuário do Office. O exemplo cria uma tabela e, quando o usuário move o foco dentro da tabela, a guia personalizada é exibida. Quando o usuário sai da tabela, a guia personalizada fica oculta. |
| [Use os atalhos do teclado para ações do suplemento do Office](/samples/officedev/pnp-officeaddins/office-add-in-keyboard-shortcuts) | Configure um projeto de suplemento básico do Excel que utiliza atalhos de teclado. |
| [Exemplo de função personalizada usando web worker](/samples/officedev/pnp-officeaddins/excel-custom-function-web-worker-pattern/) | Use web workers em funções personalizadas para evitar o bloqueio da interface do usuário do suplemento do Office. |
| [Use técnicas de armazenamento para acessar dados de um suplemento do Office quando estiver offline](/samples/officedev/pnp-officeaddins/use-storage-techniques-to-access-data-from-an-office-add-in-when-offline/) | Implemente o localStorage para habilitar a funcionalidade limitada do Suplemento do Office quando um usuário perder a conexão. |
| [Padrão de lote de função personalizada](/samples/officedev/pnp-officeaddins/excel-custom-function-batching-pattern/)| Agrupe várias chamadas em uma única chamada para reduzir o número de chamadas de rede para um serviço remoto.|

## <a name="shared-javascript-runtime"></a>Tempo de execução de JavaScript compartilhado

| Nome                | Descrição         |
|:--------------------|:--------------------|
[Compartilhe dados globais com um tempo de execução compartilhado](/samples/officedev/pnp-officeaddins/office-add-in-shared-runtime-global-data/) | Configure um projeto básico que usa o tempo de execução compartilhado para executar código para botões da faixa de opções, painel de tarefas e funções personalizadas em um único tempo de execução do navegador. |
| [Gerencie a faixa de opções e a interface do usuário do painel de tarefas e execute o código no documento aberto](/samples/officedev/pnp-officeaddins/office-add-in-ribbon-task-pane-ui/) | Crie os botões contextuais da faixa de opções que são ativados com base no estado do seu suplemento. |

## <a name="authentication-authorization-and-single-sign-on-sso"></a>Autenticação, autorização e logon único (SSO)

| Nome                | Descrição         |
|:--------------------|:--------------------|
| [Suplemento de amostra do Outlook de logon único (SSO)](/samples/officedev/pnp-officeaddins/outlook-add-in-sso-aspnet/) | Use o recurso SSO do Office para fornecer ao suplemento acesso aos dados do Microsoft Graph.|
| [Obtenha dados do OneDrive usando Microsoft Graph e msal.js em um suplemento do Office](/samples/officedev/pnp-officeaddins/office-add-in-auth-graph-react/) | Crie um suplemento do Office, como um aplicativo de página única (SPA) sem back-end, que se conecta ao Microsoft Graph e acesse pastas de trabalho armazenadas no OneDrive for Business para atualizar uma planilha.  |
| [Autenticação do suplemento do Office para o Microsoft Graph](/samples/officedev/pnp-officeaddins/office-add-in-auth-aspnet-graph/) | Aprenda a criar um suplemento do Microsoft Office que se conecte ao Microsoft Graph e acesse pastas de trabalho armazenadas no OneDrive for Business para atualizar uma planilha. |
| [Autenticação do suplemento do Outlook para Microsoft Graph](/samples/officedev/pnp-officeaddins/outlook-add-in-auth-aspnet-graph/). | Crie um suplemento do Outlook que se conecte ao Microsoft Graph e acesse pastas de trabalho armazenadas no OneDrive for Business para redigir uma nova mensagem de email. |
| [Suplemento do Office de Logon único (SSO) com ASP.NET](/samples/officedev/pnp-officeaddins/office-add-in-sso-aspnet/) | Use a API `getAccessToken` em Office.js para dar ao suplemento acesso aos dados do Microsoft Graph. Este exemplo é criado no ASP.NET. |
| [Suplemento Office dee Logon único (SSO) com Node.js](/samples/officedev/pnp-officeaddins/office-add-in-sso-nodejs/) | Use a API `getAccessToken` em Office.js para dar ao suplemento acesso aos dados do Microsoft Graph. Este exemplo é construído em Node.js.|

## <a name="additional-samples"></a>Amostras adicionais

| Nome                | Descrição         |
|:--------------------|:--------------------|
|[Use uma biblioteca compartilhada para migrar seu suplemento do Visual Studio Tools para Office para um suplemento da web do Office](/samples/officedev/pnp-officeaddins/vsto-shared-library-excel/) |Fornece uma estratégia para reutilização de código ao migrar de suplementos do VSTO para suplementos do Office. |
| [Integre uma função do Azure à sua função personalizada do Excel](/samples/officedev/pnp-officeaddins/azure-function-with-excel-custom-function/) | Integre funções do Azure com funções personalizadas para mover para a nuvem ou integrar serviços adicionais. |
|[Amostras de código DPI dinâmico](/samples/officedev/pnp-officeaddins/dynamic-dpi-code-samples/) |Uma coleção de amostras para lidar com alterações de DPI em suplementos COM, VSTO e Office. |

## <a name="next-steps"></a>Próximos passos

Ingresse no Programa Para Desenvolvedores do Microsoft 365. Obtenha uma área restrita, ferramentas e outros recursos gratuitos de que você precisa para criar soluções para a plataforma Microsoft 365.

- [A área restrita de desenvolvedor grátis](https://developer.microsoft.com/microsoft-365/dev-program#Subscription) Obtenha uma assinatura de desenvolvedor Microsoft 365 E5 gratuita e renovável por 90 dias.
- [Amostra de pacotes de dados](https://developer.microsoft.com/microsoft-365/dev-program#Sample) Configure automaticamente sua área restrita instalando dados de usuário e conteúdo para ajudá-lo a construir suas soluções.
- [Acesso a especialistas](https://developer.microsoft.com/microsoft-365/dev-program#Experts) Acesse eventos da comunidade para aprender com especialistas do Microsoft 365.
- [Recomendações personalizadas](https://developer.microsoft.com/microsoft-365/dev-program#Recommendations) Localize os recursos para desenvolvedores rapidamente em seu painel personalizado.

---
title: Vis?o geral da plataforma de Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: f0f20371eee759a449773effaff1ce365e32bf48
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/25/2018
---
# <a name="office-add-ins-platform-overview"></a>Vis?o geral da plataforma de Suplementos do Office

Voc? pode usar a plataforma de suplementos do Office para criar solu??es que estendem os aplicativos do Office e interagem com conte?do nos documentos do Office. Com os suplementos do Office, voc? pode usar tecnologias web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com o Word, Excel, PowerPoint, OneNote, Project e Outlook. Sua solu??o pode ser executada no Office atrav?s de v?rias plataformas, incluindo Office para Windows, Office Online, Office para Mac e Office para iPad.

Os suplementos do Office podem fazer quase tudo que uma p?gina da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:

-  **Adicionar novas funcionalidades para os clientes do Office** ? trazer dados externos para o Office, automatizar documentos do Office, expor a funcionalidade de terceiros em clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar aos dados que orientam a produtividade. 
    
-  **Crie novos objetos avan?ados e interativos que podem ser integrados em documentos do Office** ? Mapas, gr?ficos e visualiza??es interativas integrados que os usu?rios podem adicionar a suas pr?prias planilhas do Excel e apresenta??es do PowerPoint. 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a>Quais s?o as diferen?as entre os suplementos do Office e os suplementos de COM e VSTO? 

Os suplementos de COM ou VSTO s?o solu??es de integra??o anteriores do Office que s?o executadas apenas no Office para Windows. Ao contr?rio de suplementos de COM, os suplementos do Office n?o envolvem c?digo executado no dispositivo do usu?rio ou no cliente do Office. Para um suplemento Office, o aplicativo do host, por exemplo, o Excel, l? o manifesto do suplemento e conecta os comandos do menu e os bot?es da faixa de op??es personalizada do suplemento ? interface de usu?rio. Quando necess?rio, ele carrega o c?digo de HTML e o JavaScript, que s?o executados no contexto de um navegador em uma ?rea restrita. 

Os suplementos do Office fornecem as seguintes vantagens em rela??o aos suplementos criados usando o VBA, COM ou VSTO: 

- Suporte ? plataforma cruzada. Os suplementos do Office podem ser executados no Office para Windows, Mac, iOS e Office Online. 

- SSO (logon ?nico). Os suplementos do Office integram-se facilmente com contas do Office 365 dos usu?rios. 

- Implanta??o e distribui??o centralizada. Os administradores podem implantar suplementos do Office centralmente em uma organiza??o. 

- Acesso f?cil atrav?s da AppSource. Voc? pode disponibilizar sua solu??o para um p?blico amplo ao envi?-la para o AppSource. 

- Com base na tecnologia de Internet padr?o. Voc? pode usar qualquer biblioteca que gosta para criar suplementos do Office. 

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office 

Um suplemento do Office inclui dois componentes b?sicos: um arquivo de manifesto XML e seu pr?prio aplicativo Web. O manifesto define v?rias configura??es, incluindo como o suplemento ? integrado a clientes do Office. O aplicativo Web deve ser hospedado em um servidor Web ou servi?o de hospedagem na Web, como o Microsoft Azure.

*Figura 1. Manifesto + p?gina da Web = um Suplemento do Office*

![Manifesto mais p?gina da Web ? igual a suplemento do Office](../images/dk2-agave-overview-01.png)

### <a name="manifest"></a>Manifesto 

O manifesto ? um arquivo XML que especifica configura??es e recursos do suplemento, como os seguintes: 

- O nome de exibi??o, a descri??o, a ID, a vers?o e a localidade padr?o do suplemento. 

- Como o suplemento se integra ao Office.  

- O n?vel de permiss?o e os requisitos de acesso a dados para o suplemento. 

### <a name="web-app"></a>Aplicativo Web 

O Suplemento do Office mais b?sico consiste em uma p?gina HTML est?tica que ? exibida dentro de um aplicativo do Office, mas n?o interage com o documento do Office nem com qualquer outro recurso de Internet. No entanto, para criar uma experi?ncia que interaja com os documentos do Office ou permita que o usu?rio interaja com os recursos online de um aplicativo de host do Office, voc? pode usar qualquer tecnologia, tanto do lado do cliente como do servidor, a qual seu provedor de hospedagem d? suporte (como ASP.NET, PHP ou N?.js). Para interagir com clientes e documentos do Office, voc? usa as APIs Office.js e JavaScript. 

*Figura 2. Componentes de um suplemento Hello World do Office*

![Componentes de um suplemento Hello World](../images/dk2-agave-overview-07.png)

## <a name="extending-and-interacting-with-office-clients"></a>Estender os clientes do Office e interagir com eles 

Os suplementos do Office podem fazer o seguinte em um aplicativo de host do Office: 

-  Estender a funcionalidade (qualquer aplicativo do Office) 

-  Criar novos objetos (Excel ou PowerPoint) 
 
### <a name="extend-office-functionality"></a>Estender a funcionalidade do Office 

Voc? pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:  

-  Bot?es de faixa de op??es e comandos de menu personalizados (coletivamente chamados "comandos de suplemento") 

-  Pain?is de tarefas inser?veis 

Pain?is personalizados de interface do usu?rio e de tarefa s?o especificados no manifesto do suplemento.  

#### <a name="custom-buttons-and-menu-commands"></a>Bot?es e comandos de menu personalizados  

Voc? pode adicionar itens de menu e bot?es da faixa de op??es personalizados ? faixa de op??es, tanto no Office para ?rea de Trabalho do Windows quanto no Office Online. Isso facilita aos usu?rios o acesso ao suplemento diretamente do aplicativo do Office. Bot?es de comando podem iniciar diferentes a??es, como mostrar um painel de tarefas com c?digo HTML personalizado ou executar uma fun??o JavaScript.  

*Figura 3. Comandos do suplemento em execu??o na ?rea de Trabalho do Excel*

![Bot?es e comandos de menu personalizados](../images/add-in-commands-overview.png)

#### <a name="task-panes"></a>Pain?is de tarefas  

Voc? pode usar pain?is de tarefas, al?m dos comandos de suplemento, para permitir que os usu?rios interajam com sua solu??o. Os clientes que n?o d?o suporte aos comandos de suplemento (Office 2013 e Office para iPad) executar?o seu suplemento como um painel de tarefas. Os usu?rios iniciam os suplementos do painel de tarefas atrav?s do bot?o **Meus suplementos** na guia **Inserir**. 

*Figura 4. Painel de tarefas*

![Painel de tarefas](../images/task-pane-overview.jpg)

### <a name="extend-outlook-functionality"></a>Estender a funcionalidade do Outlook 

Os suplementos do Outlook podem estender a faixa de op??es do Office e tamb?m ser exibidos contextualmente ao lado de um item do Outlook quando voc? o exibe ou redige. Eles podem trabalhar com uma mensagem de email, uma solicita??o de reuni?o, uma resposta de reuni?o, um cancelamento de reuni?o ou um compromisso quando um usu?rio est? visualizando um item recebido, ou respondendo ou criando um novo item. 

Os suplementos do Outlook podem acessar informa??o contextual do item, como o endere?o ou a ID de rastreamento, e, em seguida, usar estes dados para acessarem informa??es adicionais sobre o servidor e de servi?os da Web para criar experi?ncias do usu?rio envolventes. Na maioria dos casos, um suplemento do Outlook ? executado sem modifica??o nos v?rios aplicativos host com suporte, incluindo Outlook, Outlook para Mac, Outlook Web App e Outlook Web App para Dispositivos para fornecer uma experi?ncia perfeita na ?rea de trabalho, na Web e em tablets e dispositivos m?veis. 

Confira a vis?o geral dos suplementos do Outlook em [Vis?o geral dos suplementos do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/). 

### <a name="create-new-objects-in-office-documents"></a>Criar novos objetos nos documentos do Office 

Voc? pode inserir objetos baseados na web, chamados de suplementos de conte?do, em documentos do Excel e PowerPoint. Com os suplementos de conte?do, voc? pode integrar visualiza??es de dados avan?adas e baseadas na Web, m?dia (como um player de v?deo do YouTube ou uma galeria de imagens) e outros tipos de conte?do externo.

*Figura 5. Suplemento de conte?do*

![Suplemento de conte?do](../images/dk2-agave-overview-05.png)

## <a name="office-javascript-apis"></a>APIs JavaScript para Office 

As APIs JavaScript para Office cont?m objetos e membros para a cria??o de suplementos e a intera??o com conte?do do Office e servi?os Web. Existe um modelo de objeto comum compartilhado pelo Excel, Outlook, Word, PowerPoint, OneNote e Project. Tamb?m existem modelos de objeto espec?ficos de host mais extensos para o Excel e o Word. Essas APIs fornecem acesso a objetos conhecidos, como par?grafos e pastas de trabalho, o que facilita a cria??o de um suplemento para um host espec?fico.  

## <a name="next-steps"></a>Pr?ximas etapas 

Para saber mais sobre como come?ar a criar o seu Suplemento do Office, experimente o nosso [In?cios R?pidos de 5 minutos](https://docs.microsoft.com/en-us/office/dev/add-ins/). Voc? pode come?ar a criar suplementos imediatamente usando o Visual Studio ou qualquer outro editor. 

Para come?ar a planejar solu??es que criem experi?ncias de usu?rio eficazes e atraentes, familiarize-se com as [diretrizes de design](../design/add-in-design.md) e as [pr?ticas recomendadas](../concepts/add-in-development-best-practices.md) para suplementos do Office.    
   
## <a name="see-also"></a>Confira tamb?m

- [Exemplos de suplementos do Office](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples)
- [No??es b?sicas da API JavaScript para Office](../develop/understanding-the-javascript-api-for-office.md)
- [Disponibilidade de host e plataforma para suplementos do Office](../overview/office-add-in-availability.md)


    

---
title: Visão geral da plataforma de Suplementos do Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: f7b1f4add776f1971e9762c5cb80dabed45b0a1c
ms.sourcegitcommit: a0e0416289b293863b8b4d3f9a12581a9e681b27
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2018
ms.locfileid: "20023162"
---
# <a name="office-add-ins-platform-overview"></a>Visão geral da plataforma de Suplementos do Office

Você pode usar a plataforma de suplementos do Office para criar soluções que estendem os aplicativos do Office e interagem com conteúdo nos documentos do Office. Com os suplementos do Office, você pode usar tecnologias web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com o Word, Excel, PowerPoint, OneNote, Project e Outlook. Sua solução pode ser executada no Office através de várias plataformas, incluindo Office para Windows, Office Online, Office para Mac e Office para iPad.

Os suplementos do Office podem fazer quase tudo que uma página da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:

-  **Adicionar novas funcionalidades para os clientes do Office** – trazer dados externos para o Office, automatizar documentos do Office, expor a funcionalidade de terceiros em clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar aos dados que orientam a produtividade. 
    
-  **Crie novos objetos avançados e interativos que podem ser integrados em documentos do Office** ‒ Mapas, gráficos e visualizações interativas integrados que os usuários podem adicionar a suas próprias planilhas do Excel e apresentações do PowerPoint. 
    
## <a name="how-are-office-add-ins-different-than-com-and-vsto-add-ins"></a>Quais são as diferenças entre os suplementos do Office e os suplementos de COM e VSTO? 

Os suplementos de COM ou VSTO são soluções de integração anteriores do Office que são executadas apenas no Office para Windows. Ao contrário de suplementos de COM, os suplementos do Office não envolvem código executado no dispositivo do usuário ou no cliente do Office. Para um suplemento Office, o aplicativo do host, por exemplo, o Excel, lê o manifesto do suplemento e conecta os comandos do menu e os botões da faixa de opções personalizada do suplemento à interface de usuário. Quando necessário, ele carrega o código de HTML e o JavaScript, que são executados no contexto de um navegador em uma área restrita. 

Os suplementos do Office fornecem as seguintes vantagens em relação aos suplementos criados usando o VBA, COM ou VSTO: 

- Suporte à plataforma cruzada. Os suplementos do Office podem ser executados no Office para Windows, Mac, iOS e Office Online. 

- SSO (logon único). Os suplementos do Office integram-se facilmente com contas do Office 365 dos usuários. 

- Implantação e distribuição centralizada. Os administradores podem implantar suplementos do Office centralmente em uma organização. 

- Acesso fácil através da AppSource. Você pode disponibilizar sua solução para um público amplo ao enviá-la para o AppSource. 

- Com base na tecnologia de Internet padrão. Você pode usar qualquer biblioteca que gosta para criar suplementos do Office. 

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office 

Um suplemento do Office inclui dois componentes básicos: um arquivo de manifesto XML e seu próprio aplicativo Web. O manifesto define várias configurações, incluindo como o suplemento é integrado a clientes do Office. O aplicativo Web deve ser hospedado em um servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure.

*Figura 1. Manifesto do suplemento (XML) + página da Web (HTML, JS) = um suplemento do Office*

![Manifesto mais página da Web é igual a suplemento do Office](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a>Manifesto 

O manifesto é um arquivo XML que especifica configurações e recursos do suplemento, como os seguintes: 

- O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento. 

- Como o suplemento se integra ao Office.  

- O nível de permissão e os requisitos de acesso a dados para o suplemento. 

### <a name="web-app"></a>Aplicativo Web 

O Suplemento do Office mais básico consiste em uma página HTML estática que é exibida dentro de um aplicativo do Office, mas não interage com o documento do Office nem com qualquer outro recurso de Internet. No entanto, para criar uma experiência que interaja com os documentos do Office ou permita que o usuário interaja com os recursos online de um aplicativo de host do Office, você pode usar qualquer tecnologia, tanto do lado do cliente como do servidor, a qual seu provedor de hospedagem dá suporte (como ASP.NET, PHP ou Nó.js). Para interagir com clientes e documentos do Office, você usa as APIs Office.js e JavaScript. 

*Figura 2. Componentes de um suplemento Hello World do Office*

![Componentes de um suplemento Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>Estender os clientes do Office e interagir com eles 

Os suplementos do Office podem fazer o seguinte em um aplicativo de host do Office: 

-  Estender a funcionalidade (qualquer aplicativo do Office) 

-  Criar novos objetos (Excel ou PowerPoint) 
 
### <a name="extend-office-functionality"></a>Estender a funcionalidade do Office 

Você pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:  

-  Botões de faixa de opções e comandos de menu personalizados (coletivamente chamados "comandos de suplemento") 

-  Painéis de tarefas inseríveis 

Painéis personalizados de interface do usuário e de tarefa são especificados no manifesto do suplemento.  

#### <a name="custom-buttons-and-menu-commands"></a>Botões e comandos de menu personalizados  

Você pode adicionar itens de menu e botões da faixa de opções personalizados à faixa de opções, tanto no Office para Área de Trabalho do Windows quanto no Office Online. Isso facilita aos usuários o acesso ao suplemento diretamente do aplicativo do Office. Botões de comando podem iniciar diferentes ações, como mostrar um painel de tarefas com código HTML personalizado ou executar uma função JavaScript.  

*Figura 3. Comandos do suplemento na faixa de opções*

![Botões e comandos de menu personalizados](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>Painéis de tarefas  

Você pode usar painéis de tarefas, além dos comandos de suplemento, para permitir que os usuários interajam com sua solução. Os clientes que não dão suporte aos comandos de suplemento (Office 2013 e Office para iPad) executarão seu suplemento como um painel de tarefas. Os usuários iniciam os suplementos do painel de tarefas através do botão **Meus suplementos** na guia **Inserir**. 

*Figura 4. Painel de tarefas*

![Painel de tarefas](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>Estender a funcionalidade do Outlook 

Os suplementos do Outlook podem estender a faixa de opções do Office e também ser exibidos contextualmente ao lado de um item do Outlook quando você o exibe ou redige. Eles podem trabalhar com uma mensagem de email, uma solicitação de reunião, uma resposta de reunião, um cancelamento de reunião ou um compromisso quando um usuário está visualizando um item recebido, ou respondendo ou criando um novo item. 

Os suplementos do Outlook podem acessar informação contextual do item, como o endereço ou a ID de rastreamento, e, em seguida, usar estes dados para acessarem informações adicionais sobre o servidor e de serviços da Web para criar experiências do usuário envolventes. Na maioria dos casos, um suplemento do Outlook é executado sem modificação nos vários aplicativos host com suporte, incluindo Outlook, Outlook para Mac, Outlook Web App e Outlook Web App para Dispositivos para fornecer uma experiência perfeita na área de trabalho, na Web e em tablets e dispositivos móveis. 

Confira a visão geral dos suplementos do Outlook em [Visão geral dos suplementos do Outlook](https://docs.microsoft.com/en-us/outlook/add-ins/). 

### <a name="create-new-objects-in-office-documents"></a>Criar novos objetos nos documentos do Office 

Você pode inserir objetos baseados na web, chamados de suplementos de conteúdo, em documentos do Excel e PowerPoint. Com os suplementos de conteúdo, você pode integrar visualizações de dados avançadas e baseadas na Web, mídia (como um player de vídeo do YouTube ou uma galeria de imagens) e outros tipos de conteúdo externo.

*Figura 5. Suplemento de conteúdo*

![Suplemento de conteúdo](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>APIs JavaScript para Office 

As APIs JavaScript para Office contêm objetos e membros para a criação de suplementos e a interação com conteúdo do Office e serviços Web. Existe um modelo de objeto comum compartilhado pelo Excel, Outlook, Word, PowerPoint, OneNote e Project. Também existem modelos de objeto específicos de host mais extensos para o Excel e o Word. Essas APIs fornecem acesso a objetos conhecidos, como parágrafos e pastas de trabalho, o que facilita a criação de um suplemento para um host específico.  

## <a name="next-steps"></a>Próximas etapas 

Para saber mais sobre como começar a compilar seu suplemento do Office, experimente o nosso [Inícios Rápidos de 5 minutos](https://docs.microsoft.com/en-us/office/dev/add-ins/). Você pode começar a compilar suplementos imediatamente usando o Visual Studio ou qualquer outro editor. 

Para começar a planejar soluções que criem experiências de usuário eficazes e atraentes, familiarize-se com as [diretrizes de design](../design/add-in-design.md) e as [práticas recomendadas](../concepts/add-in-development-best-practices.md) para suplementos do Office.    
   
## <a name="see-also"></a>Confira também

- [Exemplos de suplementos do Office](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples)
- [Noções básicas da API JavaScript para Office](../develop/understanding-the-javascript-api-for-office.md)
- [Disponibilidade de host e plataforma para suplementos do Office](../overview/office-add-in-availability.md)


    

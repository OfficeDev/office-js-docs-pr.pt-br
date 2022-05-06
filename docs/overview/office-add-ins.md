---
title: Visão geral da plataforma de Suplementos do Office
description: Use tecnologias da Web conhecidas, como HTML, CSS e JavaScript, para estender e interagir com os aplicativos Word, Excel, PowerPoint, OneNote, Project e Outlook.
ms.date: 04/14/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 8216abbce1147280c722b2ac8450379e2425172b
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244853"
---
# <a name="office-add-ins-platform-overview"></a>Visão geral da plataforma de Suplementos do Office

Use a plataforma de Suplementos do Office para criar soluções que estendem os aplicativos do Office e interagem com o conteúdo dos documentos do Office. Com os Suplementos do Office, você pode usar tecnologias da Web conhecidas, como HTML, CSS e JavaScript para estender e interagir com o Outlook, Excel, Word, PowerPoint, OneNote e Project. Sua solução pode ser executada no Office em várias plataformas, incluindo no Windows, Mac, iPad e em um navegador.

![O aplicativo do Office mais um site inserido (suplemento) tornam infinitas as possibilidades de extensibilidade.](../images/addins-overview.png)

Os suplementos do Office podem fazer quase tudo que uma página da Web pode fazer dentro do navegador. Use a plataforma de suplementos do Office para:

- **Adicione novas funcionalidades aos clientes do Office** - Traga dados externos para o Office, automatize documentos do Office, exponha funcionalidades da Microsoft e de outros clientes do Office e muito mais. Por exemplo, use a API do Microsoft Graph para se conectar a dados que impulsionam a produtividade.

- **Crie novos objetos avançados e interativos que podem ser integrados em documentos do Office** ‒ Mapas, gráficos e visualizações interativas integrados que os usuários podem adicionar a suas próprias planilhas do Excel e apresentações do PowerPoint.

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a>Quais são as diferenças entre os suplementos do Office e os suplementos de COM e VSTO?

Os suplementos de COM ou VSTO são soluções de integração anteriores do Office que são executadas apenas no Office no Windows. Ao contrário de suplementos de COM, os suplementos do Office não envolvem código executado no dispositivo do usuário ou no cliente do Office. Para um suplemento do Office, o aplicativo do cliente (por exemplo, o Excel), lê o manifesto do suplemento e conecta os comandos do menu e os botões da faixa de opções personalizada do suplemento à interface de usuário. Quando necessário, ele carrega o código de HTML e o JavaScript, que são executados no contexto de um navegador em uma área restrita.

![Os motivos para usar os Suplementos do Office: multiplataforma, implantação centralizada, acesso fácil por meio do AppSource e baseado em tecnologias Web padrão.](../images/why.png)

Os Suplementos do Office oferecem as seguintes vantagens em relação aos suplementos criados usando VBA, COM ou VSTO.

- Suporte à plataforma cruzada. Os suplementos do Office podem ser executados no Office na Web, Windows, Mac e iPad.

- Implantação e distribuição centralizadas. Os administradores podem implantar suplementos do Office centralmente em uma organização.

- Acesso fácil através da AppSource. Você pode disponibilizar sua solução para um público amplo ao enviá-la para o AppSource.

- Com base na tecnologia de Internet padrão. Você pode usar qualquer biblioteca que gosta para criar suplementos do Office.

## <a name="components-of-an-office-add-in"></a>Componentes de um suplemento do Office

Um suplemento do Office inclui dois componentes básicos: um arquivo de manifesto XML e seu próprio aplicativo Web. O manifesto define várias configurações, incluindo como o suplemento é integrado a clientes do Office. O aplicativo Web deve ser hospedado em um servidor Web ou serviço de hospedagem na Web, como o Microsoft Azure.

### <a name="manifest"></a>Manifesto

O manifesto é um arquivo XML que especifica configurações e recursos do suplemento, como os seguintes:

- O nome de exibição, a descrição, a ID, a versão e a localidade padrão do suplemento.

- Como o suplemento se integra ao Office.  

- O nível de permissão e os requisitos de acesso a dados para o suplemento.

### <a name="web-app"></a>Aplicativo Web

O Suplemento do Office mais básico consiste em uma página HTML estática que é exibida dentro de um aplicativo do Office, mas não interage com o documento do Office nem com qualquer outro recurso de Internet. No entanto, para criar uma experiência que interaja com os documentos do Office ou permita que o usuário interaja com os recursos online de um aplicativo cliente do Office, você pode usar qualquer tecnologia, tanto do lado do cliente como do servidor, a qual seu provedor de hospedagem dá suporte (como ASP.NET, PHP ou Node.js). Para interagir com clientes e documentos do Office, você usa as APIs Office.js e JavaScript.

![Componentes de um suplemento Hello World.](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>Estender os clientes do Office e interagir com eles

Os Suplementos do Office podem fazer o seguinte em um aplicativo cliente do Office.

- Estender a funcionalidade (qualquer aplicativo do Office)

- Criar novos objetos (Excel ou PowerPoint)

### <a name="extend-office-functionality"></a>Estender a funcionalidade do Office

Você pode adicionar novas funcionalidades a aplicativos do Office por meio do seguinte:  

- Botões de faixa de opções e comandos de menu personalizados (coletivamente chamados "comandos de suplemento")

- Painéis de tarefas inseríveis

Painéis personalizados de interface do usuário e de tarefa são especificados no manifesto do suplemento.  

#### <a name="custom-buttons-and-menu-commands"></a>Botões e comandos de menu personalizados  

Você pode adicionar itens de menu e botões da faixa de opções personalizados à faixa de opções do Office na Web e no Windows. Isso facilita o acesso dos usuários ao seu suplemento diretamente no aplicativo do Office. Os botões de comando podem iniciar ações diferentes, como mostrar um painel de tarefas com HTML personalizado ou executar uma função JavaScript.  

![Botões e comandos de menu personalizados.](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>Painéis de tarefas  

Você pode usar painéis de tarefas, além dos comandos de suplemento, para permitir que os usuários interajam com sua solução. Os clientes que não dão suporte aos comandos de suplemento (Office 2013 e Office para iPad) executarão seu suplemento como um painel de tarefas. Os usuários iniciam os suplementos do painel de tarefas através do botão **Meus suplementos** na guia **Inserir**.

![Usar painéis de tarefas, além dos comandos do suplemento.](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>Estender a funcionalidade do Outlook

Os suplementos do Outlook podem estender a faixa de opções do aplicativo do Office e também serem exibidos contextualmente ao lado de um item do Outlook quando você o exibe ou redige. Eles podem funcionar com uma mensagem de email, uma solicitação de reunião, uma resposta de reunião, um cancelamento de reunião ou um compromisso quando um usuário estiver visualizando um item recebido, respondendo ou criando um novo item.

Os suplementos do Outlook podem acessar informações contextuais do item, como o endereço ou a ID de rastreamento, e usar esses dados para acessar informações adicionais no servidor e de serviços Web para criar experiências de usuário envolventes. Na maioria dos casos, um suplemento do Outlook é executado sem modificação no aplicativo cliente do Outlook para fornecer uma experiência perfeita na área de trabalho, na Web e em dispositivos móveis e tablet.

Confira a visão geral dos suplementos do Outlook em [Visão geral dos suplementos do Outlook](../outlook/outlook-add-ins-overview.md).

### <a name="create-new-objects-in-office-documents"></a>Criar novos objetos nos documentos do Office

Você pode inserir objetos baseados na web, chamados de suplementos de conteúdo, em documentos do Excel e PowerPoint. Com os suplementos de conteúdo, você pode integrar visualizações de dados avançadas e baseadas na Web, mídia (como um player de vídeo do YouTube ou uma galeria de imagens) e outros tipos de conteúdo externo.

![Inserir objetos baseados na Web chamados suplementos de conteúdo.](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>APIs JavaScript para Office

As APIs JavaScript para Office contêm objetos e membros para a criação de suplementos e a interação com conteúdo do Office e serviços Web. Existe um modelo de objeto comum compartilhado pelo Excel, Outlook, Word, PowerPoint, OneNote e Project. Também existem modelos de objeto específicos de aplicativo mais extensos para o Excel e o Word. Essas APIs fornecem acesso a objetos conhecidos, como parágrafos e pastas de trabalho, o que facilita a criação de um suplemento para um aplicativo específico.

## <a name="next-steps"></a>Próximas etapas

Para obter uma introdução mais detalhada sobre o desenvolvimento de Suplementos do Office, confira [Desenvolver suplementos do Office](../develop/develop-overview.md).

## <a name="see-also"></a>Confira também

- [Principais conceitos dos Suplementos do Office](../overview/core-concepts-office-add-ins.md)
- [Desenvolver Suplementos do Office](../develop/develop-overview.md)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)
- [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)

---
title: 'Desenvolver Suplementos do Office '
description: Uma introdução ao desenvolvimento de Suplementos do Office.
ms.date: 03/11/2022
ms.localizationpriority: high
ms.openlocfilehash: e5f053535afd852b2c71edcfa52d8b4f4a1e54dd
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743518"
---
# <a name="develop-office-add-ins"></a>Desenvolver Suplementos do Office 

> [!TIP]
> Análise a [Visão geral da plataforma de Suplementos do Office](../overview/office-add-ins.md) antes de ler este artigo.

Todos os Suplementos do Office são criados com base na plataforma de Suplementos do Office. Para qualquer suplemento que você criar, você precisará entender conceitos importantes como disponibilidade de aplicativo e plataforma, padrões de programação da API do Office JavaScript, como especificar as configurações e recursos de um suplemento no arquivo de manifesto, como projetar a Interface do Usuário, experiência e muito mais. Conceitos básicos de desenvolvimento como esses são abordados aqui na seção **Ciclo de vida de desenvolvimento** > **Desenvolver** da documentação. Análise as informações contidas aqui antes de explorar a documentação específica do aplicativo que corresponde ao suplemento que você está criando (por exemplo, [Excel](../excel/index.yml)).

## <a name="create-an-office-add-in"></a>Criar um Suplemento do Office

Você pode criar um suplemento do Office usando o [Gerador Yeoman para suplementos do Office](yeoman-generator-overview.md) ou Visual Studio.

### <a name="yeoman-generator"></a>Gerador do Yeoman

O gerador Yeoman para Suplementos do Office pode ser usado para criar um projeto de Suplemento do Office com Node.js que pode ser gerenciado com o Visual Studio Code ou qualquer outro editor. O gerador pode criar Suplementos do Office para qualquer um dos seguintes:

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Funções personalizadas do Excel

Crie seu projeto usando HTML, CSS e JavaScript (ou TypeScript) ou usando o Angular ou React. Para qualquer estrutura escolhida, você pode escolher entre o JavaScript e o Typescript também. Para saber mais sobre como criar suplementos com o gerador, confira [Gerador do Yeoman para Suplementos do Office](yeoman-generator-overview.md).

### <a name="visual-studio"></a>Visual Studio

O Visual Studio pode ser usado para criar suplementos do Office para o Excel, Outlook, Word e PowerPoint. Um projeto do suplemento do Office é criado como parte de uma solução do Visual Studio e usa HTML, CSS e JavaScript. Para saber mais sobre como criar suplementos usando o Visual Studio, confira [Desenvolver suplementos do Office com o Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understand-the-two-parts-of-an-office-add-in"></a>Entender as duas partes de um Suplemento do Office

Um suplemento do Office consiste em duas partes.

- O manifesto do suplemento (um arquivo XML) que defina as configurações e recursos do suplemento.

- O aplicativo Web que defina a interface do usuário e a funcionalidade de componentes do suplemento, como painéis de tarefas, suplementos de conteúdo e caixas de diálogo.

O aplicativo Web usa a API JavaScript para Office para interagir com o conteúdo do documento do Office no qual o suplemento está sendo executado. Seu suplemento também pode fazer outras coisas que os aplicativos Web normalmente fazem, como chamar serviços Web externos, facilitar a autenticação do usuário e mais.

### <a name="define-an-add-ins-settings-and-capabilities"></a>Definir as configurações e os recursos do suplemento

Um manifesto do suplemento do Office (um arquivo XML) define as configurações e os recursos do suplemento. Você vai configurar o manifesto para especificar itens como:

- Metadados que descrevem o suplemento (por exemplo, ID, versão, descrição, nome de exibição, local padrão).
- Aplicativos do Office onde o suplemento será executado.
- Permissões necessárias para o suplemento.
- Como o suplemento se integra ao Office, incluindo qualquer interface do usuário personalizada que o suplemento cria (por exemplo, guias personalizadas, botões da faixa de opções).
- Localização de imagens que o suplemento usa para identidade visual e iconografia de comando.
- Dimensões do suplemento (por exemplo, dimensões para suplementos de conteúdo, altura solicitada para suplementos do Outlook).
- As regras que especificam quando o suplemento é ativado no contexto de uma mensagem ou de um compromisso (somente para suplementos do Outlook).

Para saber mais sobre o manifesto, confira [Manifesto XML de suplementos do Office](add-in-manifests.md).

### <a name="interact-with-content-in-an-office-document"></a>Interagir com o conteúdo em um documento do Office

Um suplemento do Office pode usar as APIs JavaScript para Office para interagir com o conteúdo no documento do Office no qual o suplemento está sendo executado.

#### <a name="access-the-office-javascript-api-library"></a>Acessar a biblioteca de API JavaScript do Office

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a>Modelos de API

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a>Conjuntos de requisitos da API

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="explore-apis-with-script-lab"></a>Explorar as APIs com o Script Lab

O Script Lab é um suplemento que permite explorar a API JavaScript para Office e executar trechos de código enquanto você trabalha em um programa do Office, como o Excel ou o Word. Ele está disponível gratuitamente através do [AppSource](https://appsource.microsoft.com/product/office/WA104380862) e é uma ferramenta útil para incluir no seu kit de ferramentas de desenvolvimento ao prototipar e verificar a funcionalidade desejada no suplemento. No Script Lab, você pode acessar uma biblioteca de exemplos internos para experimentar APIs rapidamente ou até mesmo usar uma amostra como o ponto de partida para o seu próprio código.

O vídeo de um minuto a seguir mostra o Script Lab em ação.

[![Vídeo curto mostrando o Script Lab em execução no Excel, Word e PowerPoint.](../images/screenshot-wide-youtube.png 'Visualização de vídeo do Script Lab')](https://aka.ms/scriptlabvideo)

Para saber mais sobre o Script Lab, confira [Explorar as APIs JavaScript para Office usando o Script Lab](../overview/explore-with-script-lab.md).

## <a name="extend-the-office-ui"></a>Estender a interface do usuário do Office

Um suplemento do Office pode estender a interface do usuário do Office usando comandos de suplementos e contêineres HTML como painéis de tarefas, suplementos de conteúdo ou caixas de diálogo.

- Os [comandos de suplemento](../design/add-in-commands.md) podem ser usados para adicionar guias, botões e menus personalizados à faixa de opções padrão no office ou para estender o menu de contexto padrão que aparece quando os usuários clicam com o botão direito do mouse em um texto em um documento do Office ou em um objeto no Excel. Quando os usuários selecionam um comando de suplemento, eles iniciam a tarefa que o comando de suplemento especifica, como a execução de código JavaScript, a abertura de um painel de tarefas ou a inicialização de uma caixa de diálogo.

- Os contêineres HTML como [painéis de tarefas](../design/task-pane-add-ins.md), [suplementos de conteúdo](../design/content-add-ins.md) e [caixas de diálogo](../design/dialog-boxes.md) podem ser usadas para exibir a interface do usuário personalizada e expor uma funcionalidade adicional em um aplicativo do Office. O conteúdo e a funcionalidade de cada painel de tarefas, suplemento de conteúdo ou caixa de diálogo são derivados de uma página da Web que você especifica. Essas páginas da Web podem usar a API JavaScript para Office para interagir com o conteúdo do documento do Office no qual o suplemento está sendo executado, além disso, também pode fazer outras coisas que as páginas da Web geralmente fazem, como chamar serviços Web externos, facilitar a autenticação do usuário e mais.

A imagem a seguir mostra um comando de suplemento na faixa de opções, um painel de tarefas à direita do documento e uma caixa de diálogo ou suplemento de conteúdo sobre o documento.

![Diagrama mostrando comandos de suplemento na faixa de opções, um painel de tarefas e um suplemento de conteúdo/caixa de diálogo em um documento do Office.](../images/add-in-ui-elements.png)

Para obter mais informações sobre como estender a Interface de Usuário do Office e projetar a Experiência de Usuário do suplemento, confira [Elementos da Interface do Usuário do Office para Suplementos do Office](../design/interface-elements.md).

## <a name="next-steps"></a>Próximos passos

Este artigo descreveu as diferentes maneiras de criar suplementos do Office, apresentou as maneiras como um suplemento pode estender a IU do Office, descreveu os conjuntos de API e apresentou o Script Lab como uma ferramenta valiosa para explorar as APIs de JavaScript do Office e a funcionalidade de suplemento de criação de protótipo. Agora que você já explorou estas informações introdutórias, considere continuar sua jornada de suplementos do Office ao longo dos caminhos a seguir.

### <a name="create-an-office-add-in"></a>Criar um Suplemento do Office

Você pode criar rapidamente um suplemento básico para o Excel, o OneNote, o Outlook, o PowerPoint, o Project ou o Word realizando um [início rápido de 5 minutos](../index.yml). Se você já concluiu um início rápido e deseja criar um suplemento um pouco mais complexo, experiente o [tutorial](../index.yml).

### <a name="learn-more"></a>Saiba mais

Saiba mais sobre o desenvolvimento, testes e publicação de suplementos do Office explorando essa documentação.

> [!TIP]
> Para qualquer suplemento que você construir, você usará informações na seção [Ciclo de vida de desenvolvimento](../overview/core-concepts-office-add-ins.md)desta documentação, junto com informações na seção específica do aplicativo que corresponde ao tipo de suplemento que você está construindo (por exemplo, [Excel](../excel/index.yml)).

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](../overview/office-add-ins.md)
- [Saiba mais sobre o Programa para Desenvolvedores do Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Fazer o design de Suplementos do Office](../design/add-in-design.md)
- [Testar e depurar Suplementos do Office](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)

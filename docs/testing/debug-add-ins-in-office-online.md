---
title: Depurar suplementos no Office na Web
description: Como usar o Office na Web para testar e depurar seus suplementos.
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: c8c67be0fe35d6aa4ebe7771fb261101d58d1c3d
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128402"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Depurar suplementos no Office na Web


Você pode criar e depurar suplementos em um computador que não esteja executando o Windows ou os clientes de área de trabalho do Office 2013 ou do Office 2016, por exemplo, se você estiver desenvolvendo no Mac. Este artigo descreve como usar o Office Online para testar e depurar seus suplementos. Este artigo descreve como usar o Office na Web para testar e depurar seus suplementos. 

## <a name="prerequisites"></a>Pré-requisitos

Para começar:

- Obtenha uma conta de desenvolvedor do Office 365, se já não tiver uma, ou o acesso a um site do SharePoint.

  > [!NOTE]
  > Para se inscrever para uma assinatura gratuita de desenvolvedor do Office 365, ingresse no [Programa de Desenvolvedor do Office 365](https://developer.microsoft.com/office/dev-program). Confira as instruções passo a passo sobre como participar do Programa para Desenvolvedores do Office 365, entrar e configurar sua assinatura na [documentação do Programa para Desenvolvedores do Office 365](/office/developer-program/office-365-developer-program).

- Configure um catálogo de aplicativos no Office 365 (SharePoint Online). Um catálogo de aplicativos é um conjunto de sites dedicado no SharePoint Online, o qual hospeda bibliotecas de documentos para suplementos do Office. Se você tiver seu próprio site do SharePoint, poderá configurar uma biblioteca de documentos do catálogo de aplicativos. Para obter mais informações, consulte [Publicar suplementos de painel de tarefas e de conteúdo em um catálogo de aplicativos no SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a>Depurar seu suplemento do Excel ou Word na Web

Para depurar seu suplemento usando o Office na Web:

1. Implante o suplemento em um servidor que dê suporte a SSL.

    > [!NOTE]
    > Recomendamos que você use o [gerador Yeoman](https://github.com/OfficeDev/generator-office) para criar e hospedar seu suplemento.

2. No seu [arquivo de manifesto de suplemento](../develop/add-in-manifests.md), atualize o valor do elemento **SourceLocation** para incluir um URI absoluto, em vez de relativo. Por exemplo:

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. Carregue o manifesto para a biblioteca de suplementos do Office no catálogo de aplicativos no SharePoint.

4. Inicie o Excel ou Word na Web do inicializador de aplicativos no Office 365 e abra um novo documento.

5. Na guia Inserir, escolha **Meus Suplementos** ou **Suplementos do Office** para inserir seu suplemento e testá-lo no aplicativo.

6. Use seu depurador de navegador favorito para depurar o suplemento.

## <a name="potential-issues"></a>Possíveis problemas

A seguir apresentamos alguns problemas que você pode encontrar ao depurar:

- Alguns erros de JavaScript que você vê podem vir do Office na Web.

- O navegador pode mostrar um erro de certificado inválido que você deve ignorar. O processo para fazer isso varia com o navegador e as interfaces de usuário dos vários navegadores para fazer essa alteração periodicamente. Você deve pesquisar na ajuda do navegador ou pesquisar online para obter instruções. (Por exemplo, procure por "Aviso de certificado inválido do Microsoft Edge".) A maioria dos navegadores terá um link na página de aviso que permite que você clique na página do suplemento. Por exemplo, o Microsoft Edge possui um link "Ir para a página da Web (não recomendado)". Mas você geralmente terá que passar por este link toda vez que o suplemento for recarregado. Para um bypass mais duradouro, consulte a ajuda, como sugerido.

- Se você definir pontos de interrupção no seu código, o Office na Web pode lançar uma mensagem de erro indicando que não é possível salvar.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Políticas de validação do AppSource](/office/dev/store/validation-policies)  
- [Criar aplicativos e suplementos eficazes para o AppSource](/office/dev/store/create-effective-office-store-listings)  
- [Solucionar erros de usuários com suplementos do Office](testing-and-troubleshooting.md)
    

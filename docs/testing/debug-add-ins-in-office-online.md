---
title: Depurar suplementos no Office na Web
description: Como usar o Office na Web para testar e depurar seus suplementos.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: f7ef3fa3d6389629e28b428b9bdbe3b128896b1f
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094488"
---
# <a name="debug-add-ins-in-office-on-the-web"></a>Depurar suplementos no Office na Web

Você pode criar e depurar suplementos em um computador que não esteja executando o Windows ou os clientes de área de trabalho do Office 2013 ou do Office 2016, por exemplo, se você estiver desenvolvendo no Mac. Este artigo descreve como usar o Office Online para testar e depurar seus suplementos. Este artigo descreve como usar o Office na Web para testar e depurar seus suplementos. 

## <a name="prerequisites"></a>Pré-requisitos

Para começar:

- Obtenha uma conta de desenvolvedor do Microsoft 365 se você ainda não tiver um ou tiver acesso a um site do SharePoint.

  > [!NOTE]
  > Para obter uma assinatura de desenvolvedor do Microsoft 365, disponível em 90 dias, ingresse no [programa de desenvolvedor do microsoft 365](https://developer.microsoft.com/office/dev-program). Consulte a [documentação do programa para desenvolvedores do microsoft 365](/office/developer-program/office-365-developer-program) para obter instruções passo a passo sobre como participar do programa de desenvolvedor do Microsoft 365 e configurar sua assinatura.

- Configurar um catálogo de aplicativos no SharePoint Online. Um catálogo de aplicativos é um conjunto de sites dedicado no SharePoint Online que hospeda bibliotecas de documentos para suplementos do Office. Se você tiver seu próprio site do SharePoint, você pode configurar uma biblioteca de documentos do catálogo de aplicativos. Para obter mais informações, confira [publicar suplementos de conteúdo e painel de tarefas em um catálogo de aplicativos no SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).


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

4. Inicie o Excel ou Word na Web do inicializador de aplicativos no Microsoft 365 e abra um novo documento.

5. Na guia Inserir, escolha **meus** suplementos ou **suplementos do Office** para inserir seu suplemento e testá-lo no aplicativo.

6. Use seu depurador de navegador favorito para depurar o suplemento.

## <a name="potential-issues"></a>Possíveis problemas

A seguir apresentamos alguns problemas que você pode encontrar ao depurar:

- Alguns erros de JavaScript que você vê podem vir do Office na Web.

- O navegador pode mostrar um erro de certificado inválido que você deve ignorar. O processo para fazer isso varia com o navegador e as interfaces de usuário dos vários navegadores para fazer essa alteração periodicamente. Você deve pesquisar na ajuda do navegador ou pesquisar online para obter instruções. (Por exemplo, procure por "Aviso de certificado inválido do Microsoft Edge".) A maioria dos navegadores terá um link na página de aviso que permite que você clique na página do suplemento. Por exemplo, o Microsoft Edge possui um link "Ir para a página da Web (não recomendado)". Mas você geralmente terá que passar por este link toda vez que o suplemento for recarregado. Para um bypass mais duradouro, consulte a ajuda, como sugerido.

- Se você definir pontos de interrupção no seu código, o Office na Web pode lançar uma mensagem de erro indicando que não é possível salvar.

## <a name="see-also"></a>Confira também

- [Práticas recomendadas para o desenvolvimento de suplementos do Office](../concepts/add-in-development-best-practices.md)
- [Políticas de validação do AppSource](/legal/marketplace/certification-policies)  
- [Criar aplicativos e suplementos eficazes para o AppSource](/office/dev/store/create-effective-office-store-listings)  
- [Solucionar erros de usuários com suplementos do Office](testing-and-troubleshooting.md)

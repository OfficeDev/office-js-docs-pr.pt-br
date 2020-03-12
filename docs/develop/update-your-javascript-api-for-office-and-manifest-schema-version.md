---
title: Atualizar para a biblioteca de API JavaScript do Office mais recente e o esquema de manifesto de suplemento versão 1,1
description: Atualize seus arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 10/11/2019
localization_priority: Normal
ms.openlocfilehash: ccb6dade39850a656326482c3abc824ecfb6637e
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596582"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Atualizar para a biblioteca de API JavaScript do Office mais recente e o esquema de manifesto de suplemento versão 1,1

Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.

> [!NOTE]
> Os projetos criados no Visual Studio 2019 já usarão a versão 1,1. No entanto, há atualizações secundárias ocasionais para a versão 1.1 que você pode aplicar ao usar as técnicas neste artigo.

## <a name="use-the-most-up-to-date-project-files"></a>Usar os arquivos de projeto mais atualizados

Se você usar o Visual Studio para desenvolver seu suplemento, para usar os membros mais recentes da API da API JavaScript do Office e os [recursos do v 1.1 do manifesto do suplemento](../develop/add-in-manifests.md) (que é validado no offappmanifest-1.1. xsd), será necessário baixar o Visual Studio 2019. Para baixar o Visual Studio 2019, confira a [página IDE do Visual Studio](https://visualstudio.microsoft.com/vs/). Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.

Se estiver usando um editor de texto ou IDE que não o Visual Studio para desenvolver o suplemento, é precisa atualizar as referências à CDN para o Office.js e a versão do esquema consultada pelo manifesto do suplemento.

Para executar um suplemento desenvolvido usando recursos novos e atualizados da API do Office.js e do suplemento do manifesto, seus clientes devem estar executando o Office 2013 SP1 ou uma versão posterior de produtos locais e, quando aplicável, o SharePoint Server 2013 SP1 e produtos de servidor relacionados, o Exchange Server 2013 Service Pack 1 (SP1) ou produtos hospedados online equivalentes: Office 365, SharePoint Online e Exchange Online.

Para baixar os produtos do Office, SharePoint e Exchange SP1, consulte o seguinte:

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos da área de trabalho relacionados](https://support.microsoft.com/kb/2850036)

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos do servidor relacionados](https://support.microsoft.com/kb/2850035)

- [Descrição do Exchange Server 2013 Service Pack 1](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Atualização de um projeto de suplemento do Office criado com o Visual Studio

Para projetos criados antes do lançamento da versão v 1.1 da API JavaScript do Office e do esquema de manifesto de suplemento, você pode atualizar os arquivos de um projeto usando o **Gerenciador de pacotes do NuGet**e, em seguida, atualizar as páginas HTML do suplemento para fazer referência a eles. 

Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript do Office em seu projeto para a versão mais recente
As etapas a seguir atualizarão seus arquivos de biblioteca do Office. js para a versão mais recente. As etapas usam o Visual Studio 2019, mas são semelhantes para versões anteriores do Visual Studio.

1. No Visual Studio 2019, abra ou crie um novo projeto de **suplemento do Office** .
2. Escolha **ferramentas** > **NuGet Package Manager** > **gerenciar pacotes NuGet para solução**.
3. Escolha a guia **Atualizações**.
4. Selecione Microsoft.Office.js. Verifique se a origem do pacote é de **NuGet.org**.
5. No painel esquerdo, escolha **instalar** e concluir o processo de atualização do pacote.

Você precisará realizar algumas etapas adicionais para concluir a atualização. Na marca **Head** das páginas HTML do suplemento, comente ou exclua as referências de script do Office. js existentes e faça referência à biblioteca atualizada da API JavaScript do Office da seguinte maneira:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > O `/1/` em `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Após a atualização da versão do esquema de manifesto do suplemento para 1,1, você precisará remover os **recursos** e os elementos de **capacidade** e substituí-los pelos elementos [hosts](../reference/manifest/hosts.md) e [host](../reference/manifest/host.md) ou nos [elementos requirements](specify-office-hosts-and-api-requirements.md)e requirement.

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE

Para projetos criados antes da versão do v 1.1 da API JavaScript do Office e do esquema de manifesto de suplemento, você precisa atualizar suas páginas HTML do suplemento para fazer referência à CDN da biblioteca v 1.1 e atualizar o arquivo de manifesto do suplemento para usar o esquema v 1.1. 

O processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

Você não precisa de cópias locais dos arquivos da API JavaScript do Office (Office. js e arquivos. js específicos do aplicativo) para desenvolver o suplemento do Office (fazer referência à CDN para o Office. js baixa os arquivos necessários em tempo de execução), mas se você quiser uma cópia local dos arquivos da biblioteca, poderá usar o utilitário de `Install-Package Microsoft.Office.js` [linha de comando NuGet](https://docs.nuget.org/consume/installing-nuget) e o comando para baixá-los

> [!NOTE]
> Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).


### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript do Office em seu projeto para usar a versão mais recente

1. Abra as páginas HTML do suplemento no editor de texto ou IDE.

2. Na marca **Head** das páginas HTML do suplemento, comente ou exclua as referências de script do Office. js existentes e faça referência à biblioteca atualizada da API JavaScript do Office da seguinte maneira:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > O `/1/` na frente de `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Após a atualização da versão do esquema de manifesto do suplemento para 1,1, você precisará remover os **recursos** e os elementos de **capacidade** e substituí-los pelos elementos [hosts](../reference/manifest/hosts.md) e [host](../reference/manifest/host.md) ou nos [elementos requirements](specify-office-hosts-and-api-requirements.md)e requirement.

## <a name="see-also"></a>Confira também


- [Especificar hosts do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) ]
- [Entendendo a API JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
- [Referência de esquema para manifestos de suplementos do Office (versão 1.1)](../develop/add-in-manifests.md)

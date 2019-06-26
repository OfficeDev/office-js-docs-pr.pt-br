---
title: Atualize para a API de JavaScript mais recente da biblioteca do Office e o esquema de manifesto do suplemento da versão 1.1
description: Atualize seus arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dc3d1983d653a1b914331c9aeac1d6dae9fcc772
ms.sourcegitcommit: 6d1cb188c76c09d320025abfcc99db1b16b7e37b
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/25/2019
ms.locfileid: "35226766"
---
# <a name="update-to-the-latest-javascript-api-for-office-library-and-version-11-add-in-manifest-schema"></a>Atualize para a API de JavaScript mais recente da biblioteca do Office e o esquema de manifesto do suplemento da versão 1.1

Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.

> [!NOTE]
> Projetos criados no Visual Studio 2017 já usam a versão 1.1. No entanto, há atualizações secundárias ocasionais para a versão 1.1 que você pode aplicar ao usar as técnicas neste artigo.

## <a name="use-the-most-up-to-date-project-files"></a>Usar os arquivos de projeto mais atualizados

Se estiver usando o Visual Studio para desenvolver o suplemento, para usar os membros mais recentes da API JavaScript para Office e os [recursos da v1.1 do manifesto do suplemento](../develop/add-in-manifests.md) (que é validado em relação a offappmanifest 1.1.xsd), é preciso baixar o Visual Studio 2017. Para baixar o Visual Studio 2017, confira a [página IDE do Visual Studio](https://visualstudio.microsoft.com/vs/). Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.

Se estiver usando um editor de texto ou IDE que não o Visual Studio para desenvolver o suplemento, é precisa atualizar as referências à CDN para o Office.js e a versão do esquema consultada pelo manifesto do suplemento.

Para executar um suplemento desenvolvido usando recursos novos e atualizados da API do Office.js e do suplemento do manifesto, seus clientes devem estar executando o Office 2013 SP1 ou uma versão posterior de produtos locais e, quando aplicável, o SharePoint Server 2013 SP1 e produtos de servidor relacionados, o Exchange Server 2013 Service Pack 1 (SP1) ou produtos hospedados online equivalentes: Office 365, SharePoint Online e Exchange Online.

Para baixar os produtos do Office, SharePoint e Exchange SP1, consulte o seguinte:

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos da área de trabalho relacionados](https://support.microsoft.com/kb/2850036)

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos do servidor relacionados](https://support.microsoft.com/kb/2850035)

- [Descrição do Exchange Server 2013 Service Pack 1](https://support.microsoft.com/kb/2926248)


## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Atualização de um projeto de suplemento do Office criado com o Visual Studio

Para projetos criados antes do lançamento da v1.1 da API JavaScript para Office e o esquema de manifesto do suplemento, é possível atualizar os arquivos de um projeto usando o **NuGet Package Manager** e, em seguida, atualizar as páginas HTML do suplemento para fazer referência a eles. 

Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript para Office em seu projeto para a versão mais recente
As etapas a seguir atualizam seus arquivos da biblioteca do Office para a versão mais recente. As etapas usam o Visual Studio 2017, mas são semelhantes para o Visual Studio 2015.

1. No Visual Studio 2017, abra ou crie um novo projeto de **Suplemento do Office**.
2. Escolha **Ferramentas** > **Gerenciador de Pacotes NuGet** > **Gerenciar Pacotes Nuget para a Solução**.
3. No **Gerenciador de Pacotes NuGet**, escolha **nuget.org** como **Origem do pacote**.
4. Escolha a guia **Atualizações**.
5. Selecione Microsoft.Office.js.
6. No painel à esquerda, escolha **Atualizar** e conclua o processo de atualização do pacote.

Você precisará realizar algumas etapas adicionais para concluir a atualização. Na marca **head** das páginas HTML do suplemento, comente ou exclua quaisquer referências existentes ao script office.js e faça referência à biblioteca atualizada da API JavaScript para Office da seguinte maneira:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
  ```

   > [!NOTE] 
   > O `/1/` em `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.


### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Após atualizar a versão do esquema do manifesto do suplemento para 1.1, será preciso remover os elementos **Capabilities** e **Capability** e substituí-los pelos [Hosts](/office/dev/add-ins/reference/manifest/hosts) e elementos [Host](/office/dev/add-ins/reference/manifest/host) ou pelos [elementos Requirements e Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE

Para projetos criados antes do lançamento da v1.1 da API JavaScript para Office e o esquema de manifesto de suplemento, é preciso atualizar as páginas HTML do suplemento para fazerem referência à CDN da biblioteca v1.1 e atualizar o arquivo de manifesto do suplemento para usar a v1.1 do esquema. 

O processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

Você não precisa de cópias locais dos arquivos da API JavaScript para Office (Office.js e arquivos .js específicos do aplicativo) para desenvolver um suplemento do Office (a referência à CDN para Office.js baixa os arquivos necessários no tempo de execução). Porém, se desejar uma cópia local dos arquivos da biblioteca, pode usar o [Utilitário de Linha de Comando NuGet](https://docs.nuget.org/consume/installing-nuget) e o comando `Install-Package Microsoft.Office.js` para baixá-los.

> [!NOTE]
> Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).


### <a name="update-the-javascript-api-for-office-library-files-in-your-project-to-use-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript para Office em seu projeto para usar a versão mais recente

1. Abra as páginas HTML do suplemento no editor de texto ou IDE.

2. Na marca **head** das páginas HTML do suplemento, comente ou exclua quaisquer referências existentes ao script office.js e faça referência à biblioteca atualizada da API JavaScript para Office da seguinte maneira:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/Office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > O `/1/` na frente de `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** do elemento **OfficeApp** alterando o valor de versão para `1.1` (mantendo inalterados os atributos diferentes de **xmlns**).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Após atualizar a versão do esquema do manifesto do suplemento para 1.1, será preciso remover os elementos **Capabilities** e **Capability** e substituí-los pelos [Hosts](/office/dev/add-ins/reference/manifest/hosts) e elementos [Host](/office/dev/add-ins/reference/manifest/host) ou pelos [elementos Requirements e Requirement](specify-office-hosts-and-api-requirements.md).

## <a name="see-also"></a>Confira também

- [Especificar hosts do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) ]
- [Noções básicas da API JavaScript para Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](/office/dev/add-ins/reference/javascript-api-for-office)
- [Referência de esquema para manifestos de suplementos do Office (versão 1.1)](../develop/add-in-manifests.md)

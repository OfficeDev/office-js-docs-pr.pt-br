---
title: Atualizar para a biblioteca de API JavaScript do Office mais recente e o esquema de manifesto do suplemento versão 1.1
description: Atualize seus arquivos de JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação de manifesto de suplemento usados no seu projeto de Suplemento do Office para a versão 1.1.
ms.date: 01/14/2022
ms.localizationpriority: medium
ms.openlocfilehash: 32fcadb6a36ca540a799f8d6a5dfa671ee5e5de8
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660197"
---
# <a name="update-to-the-latest-office-javascript-api-library-and-version-11-add-in-manifest-schema"></a>Atualizar para a biblioteca de API JavaScript do Office mais recente e o esquema de manifesto do suplemento versão 1.1

Este artigo descreve como atualizar os arquivos do JavaScript (Office.js e arquivos .js específicos do aplicativo) e o arquivo de validação do manifesto do suplemento no projeto do suplemento do Office para a versão 1.1.

> [!NOTE]
> Os projetos criados no Visual Studio 2019 já usarão a versão 1.1. No entanto, há atualizações secundárias ocasionais para a versão 1.1 que você pode aplicar ao usar as técnicas neste artigo.

## <a name="use-the-most-up-to-date-project-files"></a>Usar os arquivos de projeto mais atualizados

Se você usar o Visual Studio para desenvolver seu suplemento, para usar os membros mais recentes da API JavaScript do Office e os recursos [v1.1](../develop/add-in-manifests.md) do manifesto do suplemento (que é validado em relação a offappmanifest-1.1.xsd), será necessário baixar o Visual Studio 2019. Para baixar o Visual Studio 2019, consulte a página [do IDE do Visual Studio](https://visualstudio.microsoft.com/vs/). Durante a instalação, você precisará selecionar a carga de trabalho de desenvolvimento do Office/SharePoint.

Se você usar um editor de texto ou IDE diferente do Visual Studio para desenvolver seu suplemento, precisará atualizar as referências à CDN (rede de distribuição de conteúdo) para Office.js e à versão do esquema referenciada no manifesto do suplemento.

Para executar um suplemento desenvolvido usando recursos de manifesto de suplemento e API do Office.js novos e atualizados, seus clientes devem estar executando produtos locais do Office 2013 SP1 ou versão posterior e, quando aplicável, o SharePoint Server 2013 SP1 e os produtos de servidor relacionados, o Exchange Server 2013 Service Pack 1 (SP1) ou os produtos hospedados online equivalentes: Microsoft 365, SharePoint Online e Exchange Online.

Para baixar os produtos do Office, SharePoint e Exchange SP1, consulte o seguinte:

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft Office 2013 e produtos da área de trabalho relacionados](https://support.microsoft.com/kb/2850036)

- [Lista de todas as atualizações do Service Pack 1 (SP1) para o Microsoft SharePoint Server 2013 e produtos do servidor relacionados](https://support.microsoft.com/kb/2850035)

- [Descrição do Exchange Server 2013 Service Pack 1](https://support.microsoft.com/kb/2926248)

## <a name="updating-an-office-add-in-project-created-with-visual-studio"></a>Atualização de um projeto de suplemento do Office criado com o Visual Studio

Para projetos criados antes do lançamento da v1.1 da API JavaScript do Office e do esquema de manifesto do suplemento, você pode atualizar os arquivos de um projeto usando o Gerenciador de Pacotes **NuGet** e, em seguida, atualizar as páginas HTML do suplemento para fazer referência a eles.

Observe que o processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript do Office em seu projeto para a versão mais recente

As etapas a seguir atualizarão seus Office.js de biblioteca para a versão mais recente. As etapas usam o Visual Studio 2019, mas são semelhantes para versões anteriores do Visual Studio.

1. No Visual Studio 2019, abra ou crie um novo projeto **de Suplemento do Office** .
2. Escolha **Ferramentas do** > **Gerenciador de Pacotes** >  NuGet **Gerenciar Pacotes NuGet para Solução**.
3. Escolha a guia **Atualizações**.
4. Selecione Microsoft.Office.js. Verifique se a origem do pacote **é nuget.org**.
5. No painel esquerdo, escolha Instalar **e** conclua o processo de atualização do pacote.

Você precisará realizar algumas etapas adicionais para concluir a atualização. Na **marca de** cabeçalho das páginas HTML do suplemento, comente ou exclua quaisquer referências de script office.js existentes e faça referência à biblioteca de API JavaScript do Office atualizada da seguinte maneira:

  ```html
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
  ```

   > [!NOTE]
   > O `/1/` em `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** **\<OfficeApp\>** `1.1` do elemento alterando o valor de versão para (deixando atributos diferentes do **atributo xmlns** inalterados).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Depois de atualizar a versão do esquema de manifesto do suplemento para 1.1, você precisará remover os elementos **Recursos** e Funcionalidades e substituí-los por elementos [Hosts](/javascript/api/manifest/hosts) e [Host](/javascript/api/manifest/host) ou os elementos Requisitos e  Requisitos [.](specify-office-hosts-and-api-requirements.md)

## <a name="updating-an-office-add-in-project-created-with-a-text-editor-or-other-ide"></a>Atualização de um projeto de suplemento do Office criado com um editor de texto ou outro IDE

Para projetos criados antes do lançamento da v1.1 da API JavaScript do Office e do esquema de manifesto do suplemento, você precisa atualizar as páginas HTML do suplemento para referenciar a CDN da biblioteca v1.1 e atualizar o arquivo de manifesto do suplemento para usar o esquema v1.1.

O processo de atualização é aplicado _por projeto_. Você precisará repetir o processo de atualização para cada projeto de suplemento em que deseja usar a v1.1 do Office.js e o esquema de manifesto de suplemento.

Você não precisa de cópias locais dos arquivos de API JavaScript do Office (arquivos Office.js e .js específicos do aplicativo) para desenvolver um Suplemento doOffice (referenciar a CDN para Office.js baixa os arquivos necessários em runtime), mas se você quiser uma cópia local dos arquivos de biblioteca, poderá usar o Utilitário do [NuGet Command-Line](https://docs.nuget.org/consume/installing-nuget) `Install-Package Microsoft.Office.js` e o comando para baixá-los.

> [!NOTE]
> Para obter uma cópia da XSD (Definição de esquema XML) para o manifesto do suplemento v1.1, confira a listagem em [Referência de esquema para manifestos de Suplementos do Office (v1.1)](../develop/add-in-manifests.md).

### <a name="update-the-office-javascript-api-library-files-in-your-project-to-use-the-newest-release"></a>Atualizar os arquivos da biblioteca da API JavaScript do Office em seu projeto para usar a versão mais recente

1. Abra as páginas HTML do suplemento no editor de texto ou IDE.

2. Na **marca de** cabeçalho das páginas HTML do suplemento, comente ou exclua quaisquer referências de script office.js existentes e faça referência à biblioteca de API JavaScript do Office atualizada da seguinte maneira:

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
    ```

   > [!NOTE]
   > O `/1/` na frente de `office.js` na URL de CDN especifica o uso da versão incremental mais recente na versão 1 do Office.js.

### <a name="update-the-manifest-file-in-your-project-to-use-schema-version-11"></a>Atualizar o arquivo de manifesto no projeto para usar a versão 1.1 do esquema

No arquivo de manifesto do suplemento, atualize o atributo **xmlns** **\<OfficeApp\>** `1.1` do elemento alterando o valor de versão para (deixando atributos diferentes do **atributo xmlns** inalterados).

```xml
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xsi:type="ContentApp"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1">
  
  <!-- manifest contents -->

</OfficeApp>
```

> [!NOTE]
> Depois de atualizar a versão do esquema de manifesto do suplemento para 1.1, você precisará remover os elementos **Recursos** e Funcionalidades e substituí-los por elementos [Hosts](/javascript/api/manifest/hosts) e [Host](/javascript/api/manifest/host) ou os elementos Requisitos e  Requisitos [.](specify-office-hosts-and-api-requirements.md)

## <a name="see-also"></a>Confira também

- [Especificar aplicativos do Office e requisitos de API](specify-office-hosts-and-api-requirements.md) ]
- [Entendendo a API de JavaScript do Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript para Office](../reference/javascript-api-for-office.md)
- [Referência de esquema para manifestos de suplementos do Office (versão 1.1)](../develop/add-in-manifests.md)

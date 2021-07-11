---
title: Elemento Resources no arquivo de manifesto
description: O elemento Recursos contém ícones, cadeias de caracteres e URLs para o nó VersionOverrides.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 1deacc0b93e19e5f646ca2dd74d6f89de562f21e
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348290"
---
# <a name="resources-element"></a>Elemento Resources

Contém ícones, cadeias de caracteres e URLs para o nó [VersionOverrides](versionoverrides.md). Um elemento de manifesto especifica um recurso usando a **d** do recurso. Isso ajuda a manter o tamanho do manifesto manejável, especialmente quando os recursos tiverem versões para localidades diferentes. Uma **id** deve ser exclusiva dentro do manifesto e pode ter no máximo 32 caracteres.

Cada recurso pode ter um ou mais elementos filhos **Override** para definir um recurso diferente para uma localidade específica.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Tipo  |  Descrição  |
|:-----|:-----|:-----|
|  [Imagens](#images)            |  image   |  Fornece a URL HTTPS de uma imagem para um ícone. |
|  **URLs**                |  url     |  Fornece um local para a URL HTTPS. A URL pode ter 2.048 caracteres no máximo. |
|  **ShortStrings** |  cadeia de caracteres  |  O texto para os elementos **Label** e **Title**. Cada **String** contém no máximo 125 caracteres.|
|  **LongStrings**  |  cadeia de caracteres  | O texto para atributos **Description**. Cada **String** contém no máximo 250 caracteres.|

> [!NOTE]
> Use o protocolo SSL (Secure Sockets Layer) para todas as URLs nos elementos **Image** e **Url**.

### <a name="images"></a>Imagens

Cada ícone deve ter três **elementos Images,** um para cada um dos três tamanhos obrigatórios:

- 16 x 16
- 32x32
- 80x80

Os seguintes tamanhos adicionais também são suportados, mas não são necessários.

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT]
>
> - Se essa imagem for o ícone representativo do seu complemento, consulte [Create effective listings in AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) and within Office for size and other requirements.
> - O Outlook requer a capacidade de armazenar em cache os recursos de imagem para fins de desempenho. Por esse motivo, o servidor que hospeda um recurso de imagem não deve adicionar nenhuma diretriz CACHE-CONTROL ao cabeçalho da resposta. Isso fará com que o Outlook substitua automaticamente uma imagem padrão ou genérica.

## <a name="resources-examples"></a>Exemplos de recursos

```XML
<Resources>
      <bt:Images>
        <bt:Image id="icon1_16x16" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp16-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_32x32" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp32-icon_default.png" />
        </bt:Image>
        <bt:Image id="icon1_80x80" DefaultValue="https://www.contoso.com/icon_default.png">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/ja-jp80-icon_default.png" />
        </bt:Image>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residDesktopFuncUrl" DefaultValue="https://www.contoso.com/Pages/Home.aspx">
          <bt:Override Locale="ja-jp" Value="https://www.contoso.com/Pages/Home.aspx" />
        </bt:Url>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residLabel" DefaultValue="GetData">
          <bt:Override Locale="ja-jp" Value="JA-JP-GetData" />
        </bt:String>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residToolTip" DefaultValue="Get data for your document.">
          <bt:Override Locale="ja-jp" Value="JA-JP - Get data for your document." />
        </bt:String>
      </bt:LongStrings>
    </Resources>
```

```xml
<Resources>
  <bt:Images>
    <!-- Blue icon -->
    <bt:Image id="blue-icon-16" DefaultValue="YOUR_WEB_SERVER/blue-16.png"/>
    <bt:Image id="blue-icon-32" DefaultValue="YOUR_WEB_SERVER//blue-32.png"/>
    <bt:Image id="blue-icon-80" DefaultValue="YOUR_WEB_SERVER/blue-80.png"/>
  </bt:Images>
  <bt:Urls>
    <bt:Url id="functionFile" DefaultValue="YOUR_WEB_SERVER/FunctionFile/Functions.html"/>
    <!-- other URLs -->
  </bt:Urls>
  <bt:ShortStrings>
    <bt:String id="groupLabel" DefaultValue="Add-in Demo">
      <bt:Override Locale="ar-sa" Value="<Localized text>" />
    </bt:String>
    <!-- Other short strings -->
  </bt:ShortStrings>
  <bt:LongStrings>
    <bt:String id="funcReadSuperTipDescription" DefaultValue="Gets the subject of the message or appointment.">
      <bt:Override Locale="ar-sa" Value="<Localized text>." />
    </bt:String>
    <!-- Other long strings -->
  </bt:LongStrings>
</Resources>
```

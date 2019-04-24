---
title: Elemento Resources no arquivo de manifesto
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: e29e7e36585be8fd728eb46128d7ead538ea8069
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452050"
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
Cada ícone deve ter três elementos **Images**, um para cada um dos três tamanhos obrigatórios:

- 16 x 16
- 32x32
- 80x80

Os seguintes tamanhos adicionais também têm suporte, mas não são obrigatórios:

- 20x20
- 24x24
- 40x40
- 48x48
- 64x64

> [!IMPORTANT] 
> O Outlook requer a capacidade de armazenar em cache os recursos de imagem para fins de desempenho. Por esse motivo, o servidor que hospeda um recurso de imagem não deve adicionar nenhuma diretriz CACHE-CONTROL ao cabeçalho da resposta. Isso fará com que o Outlook substitua automaticamente uma imagem padrão ou genérica.    

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

---
title: Manifesto do Teams para Suplementos do Office (versão prévia)
description: Obtenha uma visão geral da versão prévia do manifesto JSON.
ms.date: 06/15/2022
ms.localizationpriority: high
ms.openlocfilehash: 8e10d553673b2c6a67166bb8d5e30a3f655c550d
ms.sourcegitcommit: c62d087c27422db51f99ed7b14216c1acfda7fba
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/08/2022
ms.locfileid: "66689380"
---
# <a name="teams-manifest-for-office-add-ins-preview"></a>Manifesto do Teams para Suplementos do Office (versão prévia)

A Microsoft está fazendo várias melhorias na plataforma de desenvolvedor do Microsoft 365. Essas melhorias fornecem mais consistência no desenvolvimento, implantação, instalação e administração de todos os tipos de extensões do Microsoft 365, incluindo os suplementos do Office. Essas alterações são compatíveis com os suplementos existentes. 

Uma melhoria importante na qual estamos trabalhando é a capacidade de criar uma única unidade de distribuição para todas as extensões do Microsoft 365 usando o mesmo formato de manifesto e esquema, com base no manifesto atual do Teams formatado em JSON.

Demos um primeiro passo importante para essas metas, possibilitando que você crie Suplementos do Outlook, em execução somente no Windows, com uma versão do manifesto JSON do Teams.

> [!NOTE]
> O novo manifesto está disponível para visualização e está sujeito a alterações com base nos comentários. Incentivamos os desenvolvedores de suplementos experientes a experimentá-lo. O manifesto de visualização não deve ser usado em suplementos de produção. 

Durante o período de visualização inicial, as limitações a seguir se aplicam.

- A versão prévia do manifesto do Teams dá suporte apenas a suplementos do Outlook e somente à assinatura do Office para Windows. Estamos trabalhando para estender o suporte para o Excel, o PowerPoint e o Word.
- Ainda não é possível combinar e realizar sideload de um suplemento com um aplicativo do Teams, como uma guia pessoal do Teams ou outros tipos de extensão do Microsoft 365. Nos próximos meses, continuaremos a estender a versão prévia para dar suporte a esses cenários e fornecer ferramentas adicionais para atualizar manifestos para o formato de visualização prévia.

> [!TIP]
> Pronto para começar a usar a versão prévia do manifesto do Teams? Comece por [Criar um suplemento do Outlook com um manifesto do Teams (versão prévia)](../quickstarts/outlook-quickstart-json-manifest.md).

## <a name="overview-of-the-json-manifest"></a>Visão geral do manifesto JSON

### <a name="schemas-and-general-points"></a>Esquemas e pontos gerais

Há apenas um esquema para a [o manifesto JSON na versão prévia](/microsoftteams/platform/resources/dev-preview/developer-preview-intro), em comparação ao manifesto XML atual que tem um total de sete [Esquemas](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).  

### <a name="conceptual-mapping-of-the-preview-json-and-current-xml-manifests"></a>Mapeamento conceitual da versão prévia do JSON e dos manifestos XML atuais

Esta seção descreve a versão prévia do manifesto JSON para leitores que estão familiarizados com o manifesto XML atual. Alguns pontos a serem considerados: 

- O JSON não distingue entre o atributo e o valor do elemento, como o XML faz. Normalmente, o JSON que mapeia para um elemento XML torna o valor do elemento e cada um dos atributos uma propriedade filho. O exemplo a seguir mostra algumas marcações XML e seu equivalente JSON.
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- Há muitos locais no manifesto XML atual em que um elemento com um nome plural tem filhos com a versão singular do mesmo nome. Por exemplo, a marcação para configurar um menu personalizado inclui um elemento **\<Items\>** que pode ter vários elementos filhos **\<Item\>**. O equivalente JSON desses elementos plurais é uma propriedade que tem uma matriz como seu valor. Os membros da matriz são objetos *anônimos*, não propriedades chamadas "item" ou "item1", "item2", etc. O item a seguir é um exemplo.

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

#### <a name="top-level-structure"></a>Estrutura de nível superior

O nível de raiz do manifesto JSON de visualização, que corresponde aproximadamente ao elemento **\<OfficeApp\>** no manifesto XML atual, é um objeto anônimo. 

Os filhos de **\<OfficeApp\>** são geralmente divididos em duas categorias nocionais. O elemento **\<VersionOverrides\>** é uma categoria. O outro consiste de todos os outros filhos de **\<OfficeApp\>**, que são coletivamente referidos como o manifesto base. Portanto, a versão prévia do manifesto JSON também tem uma divisão semelhante. Existe uma propriedade de "extensão" de nível superior que corresponde aproximadamente em suas finalidades e propriedades filho ao elemento **\<VersionOverrides\>**. A versão prévia do manifesto JSON também tem mais de 10 outras propriedades de nível superior que atendem coletivamente às mesmas finalidades que o manifesto base do manifesto XML. Essas outras propriedades podem ser consideradas coletivamente como o manifesto base do manifesto JSON. 

> [!NOTE]
> Quando for possível combinar um suplemento com outros tipos de extensão do Microsoft 365 em um único manifesto, haverá outras propriedades de nível superior que não se encaixam na noção do manifesto base. Normalmente, haverá uma propriedade de nível superior para cada tipo de extensão do Microsoft 365, como "configurableTabs", "bots" e "conectores". Para obter exemplos, consulte a [Documentação do manifesto do Teams](/microsoftteams/platform/resources/schema/manifest-schema). Essa estrutura deixa claro que a propriedade "extensão" representa um suplemento do Office como um tipo de extensão do Microsoft 365.

#### <a name="base-manifest"></a>Manifesto base

As propriedades do manifesto base especificam características do suplemento que *qualquer* tipo de extensão do Microsoft 365 deve ter. Isso inclui guias do Teams e extensões de mensagem, não apenas suplementos do Office. Essas características incluem um nome público e uma ID exclusiva. A tabela a seguir mostra um mapeamento de algumas propriedades críticas de nível superior na versão prévia do manifesto JSON para os elementos XML no manifesto atual, em que o princípio de mapeamento *finalidade* é a finalidade da marcação.

|Propriedade JSON|Objetivo|Elemento XML|Comentários|
|:-----|:-----|:-----|:-----|
|"$schema"| Identifica o esquema do manifesto. | atributos de **\<OfficeApp\>** e **\<VersionOverrides\>** | |
|"id"| GUID do suplemento. | **\<Id\>**| |
|"versão"| A versão do suplemento. | **\<Version\>** | |
|"manifestVersion"| Versão do esquema do manifesto. |  atributos de **\<OfficeApp\>** | |
|"nome"| O nome do suplemento. | **\<DisplayName\>** | |
|"descrição"| Descrição pública do suplemento.  | **\<Description\>** | |
|"accentColor"||| Essa propriedade não tem equivalente no manifesto XML atual e não é usada na versão prévia do manifesto JSON. Mas ela deve estar presente. |
|"developer"| Identifica o desenvolvedor do suplemento. | **\<ProviderName\>** | |
|"localizationInfo"| Configura a localidade padrão e outras localidades com suporte. | **\<DefaultLocale\>** e **\<Override\>** | |
|"webApplicationInfo"| Identifica o aplicativo Web do suplemento como ele é conhecido no Azure Active Directory. | **\<WebApplicationInfo\>** | No manifesto XML atual, o elemento **\<WebApplicationInfo\>** está dentro de **\<VersionOverrides\>**, não no manifesto base. |
|"autorização"| Identifica todas as permissões do Microsoft Graph que o suplemento precisa. | **\<WebApplicationInfo\>** | No manifesto XML atual, o elemento **\<WebApplicationInfo\>** está dentro de **\<VersionOverrides\>**, não no manifesto base. |

Os elementos **\<Hosts\>**, **\<Requirements\>** e **\<ExtendedOverrides\>** fazem parte do manifesto base no manifesto XML atual. Mas conceitos e finalidades associados a esses elementos são configurados dentro da propriedade "extensão" da versão prévia do manifesto JSON. 

#### <a name="extension-property"></a>propriedade "extensão"

A propriedade "extensão" na versão prévia do manifesto JSON representa, principalmente, características do suplemento que não seriam relevantes para outros tipos de extensões do Microsoft 365. Por exemplo, os aplicativos do Office que o suplemento estende (como Excel, PowerPoint, Word e Outlook) são especificados dentro da propriedade "extensão", assim como as personalizações da faixa de opções do aplicativo do Office. As finalidades de configuração da propriedade "extension" correspondem de perto às do elemento **\<VersionOverrides\>** no manifesto XML atual.

> [!NOTE]
> A seção **\<VersionOverrides\>** do manifesto XML atual possui um sistema de "salto duplo" para muitos recursos de cadeia de caracteres. As cadeias de caracteres, incluindo URLs, são especificadas e atribuídas a uma ID no **\<Resources\>** filho de **\<VersionOverrides\>**. Elementos que exigem uma cadeia de caracteres têm um atributo `resid` que corresponde a ID de uma cadeia de caracteres no elemento **\<Resources\>**. A propriedade "extensão" da versão prévia do manifesto JSON simplifica as coisas, definindo cadeias de caracteres diretamente como valores de propriedade. Não há nada no manifesto JSON que seja equivalente ao elemento **\<Resources\>**.

A tabela a seguir mostra um mapeamento de algumas propriedades filho de alto nível da propriedade "extensão" na versão prévia do manifesto JSON para elementos XML no manifesto atual. A notação de ponto é usada para referenciar propriedades filho.

|Propriedade JSON|Objetivo|Elemento XML|Comentários|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | Identifica os conjuntos de requisitos que o suplemento precisa para ser instalado. | **\<Requirements\>** e **\<Sets\>** | |
| "requirements.scopes" | Identifica os aplicativos do Office nos quais o suplemento pode ser instalado. | **\<Hosts\>** |  |
| "faixas de opções" | As faixas de opções que o suplemento personaliza. | **\<Hosts\>**, **ExtensionPoints** e vários elementos **\*FormFactor** | As propriedade "faixas de opções" é uma matriz de objetos anônimos que mesclam as finalidades desses três elementos. Consulte [a tabela "faixas de opções"](#ribbons-table).|
| "alternativas" | Especifica a compatibilidade de versões anteriores com um suplemento COM equivalente, XLL ou ambos. | **\<EquivalentAddins\>** | Consulte [EquivalentAddins - Consulte também ](/javascript/api/manifest/equivalentaddins#see-also) para obter informações de segundo plano. |
| "runtimes"  | Configura vários tipos de suplementos que têm pouca ou nenhuma interface do usuário, como suplementos somente de função personalizada e [comandos de função](../design/add-in-commands.md#types-of-add-in-commands). | **\<Runtimes\>**. **\<FunctionFile\>** e **\<ExtensionPoint\>** (do tipo CustomFunctions) |  |
| "autoRunEvents" | Remove um manipulador de eventos de um evento especificado. | **\<Event\>** e **\<ExtensionPoint\>** (do tipo Events) |  |

##### <a name="ribbons-table"></a>tabela "faixas de opções"

A tabela a seguir mapeia as propriedades filho dos objetos filho anônimos nas "faixas de opções" matriz em elementos XML no manifesto atual. 

|Propriedade JSON|Objetivo|Elemento XML|Comentários|
|:-----|:-----|:-----|:-----|
| "contextos" | Especifica as superfícies de comando que o suplemento personaliza. | vários elementos **\*CommandSurface**, como **PrimaryCommandSurface** e **MessageReadCommandSurface** |  |
| "guias" | Configura guias personalizadas da faixa de opções. | **\<CustomTab\>** | Os nomes e a hierarquia das propriedades descendentes de "guias" correspondem de perto aos descendentes de **\<CustomTab\>**.  |

## <a name="sample-preview-json-manifest"></a>Exemplo da versão prévia do manifesto JSON

A seguir está um exemplo de uma versão prévia do manifesto JSON para um suplemento.

```json
{
  "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
  "id": "00000000-0000-0000-0000-000000000000",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Name of your app (<=30 chars)",
    "full": "Full name of app, if longer than 30 characters (<=100 chars)"
  },
  "description": {
    "short": "Short description of your app (<= 80 chars)",
    "full": "Full description of your app (<= 4000 chars)"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#230201",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "es-es",
        "file": "es-es.json"
      }
    ]
  },
  "webApplicationInfo": {
    "id": "00000000-0000-0000-0000-000000000000",
    "resource": "api://www.contoso.com/prodapp"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "Mailbox.ReadWrite.User",
          "type": "Delegated"
        }
      ]
    }
  },
  "extensions": [
    {
      "requirements": {
        "scopes": [ "mail" ],
        "capabilities": [
          {
            "name": "Mailbox", "minVersion": "1.1"
          }
        ]
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "id": "eventsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/events.html",
            "script": "https://contoso.com/events.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "onMessageSending",
              "type": "executeFunction"
            },
            {
              "id": "onNewMessageComposeCreated",
              "type": "executeFunction"
            }
          ]
        },
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.1"
              }
            ]
          },
          "id": "commandsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/commands.html",
            "script": "https://contoso.com/commands.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "action1",
              "type": "executeFunction"
            },
            {
              "id": "action2",
              "type": "executeFunction"
            },
            {
              "id": "action3",
              "type": "executeFunction"
            }
          ]
        }
      ],
      "ribbons": [
        {
          "contexts": [
            "mailCompose"
          ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    },
                    {
                      "id": "menu1",
                      "type": "menu",
                      "label": "My Menu",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "My Menu",
                        "description": "Menu with 2 actions"
                      },
                      "items": [
                        {
                          "id": "menuItem1",
                          "type": "menuItem",
                          "label": "Action 2",
                          "supertip": {
                            "title": "Action 2 Title",
                            "description": "Action 2 Description"
                          },
                          "actionId": "action2"
                        },
                        {
                          "id": "menuItem2",
                          "type": "menuItem",
                          "label": "Action 3",
                          "icons": [
                            {
                              "size": 16,
                              "file": "test_16.png"
                            },
                            {
                              "size": 32,
                              "file": "test_32.png"
                            },
                            {
                              "size": 80,
                              "file": "test_80.png"
                            }
                          ],
                          "supertip": {
                            "title": "Action 3 Title",
                            "description": "Action 3 Description"
                          },
                          "actionId": "action3"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "contexts": [ "mailRead" ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ],
      "autoRunEvents": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "events": [
            {
              "type": "newMessageComposeCreated",
              "actionId": "onNewMessageComposeCreated"
            },
            {
              "type": "messageSending",
              "actionId": "onMessageSending",
              "options": {
                "sendMode": "promptUser"
              }
            }
          ]
        }
      ],
      "alternates": [
        {
          "requirements": {
            "scopes": [ "mail" ]
          },
          "prefer": {
            "comAddin": {
              "progId": "ContosoExtension"
            }
          },
          "hide": {
            "storeOfficeAddin": {
              "officeAddinId": "00000000-0000-0000-0000-000000000000",
              "assetId": "WA000000000"
            }
          }
        }
      ]
    }
  ]
}
```

## <a name="next-steps"></a>Próximas etapas

- [Criar um suplemento do Outlook com um manifesto do Teams (versão prévia)](../quickstarts/outlook-quickstart-json-manifest.md).
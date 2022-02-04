---
title: Elemento Supertip no arquivo de manifesto
description: O elemento Supertip define uma dica de ferramenta rica (título e descrição).
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="supertip"></a>Supertip

Define uma dica de ferramenta avançada (título e descrição). É usada pelos controles de [Botão](control.md#button-control) ou [Menu](control.md#menu-dropdown-button-controls).

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
| [Title](#title) | Sim | O texto da superdica. |
| [Descrição](#description) | Sim | A descrição da superdica.<br>**Observação**: (Outlook) Somente clientes Windows e Mac são suportados. |

### <a name="title"></a>Título

Obrigatório. O texto da superdica. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no elemento **ShortStrings** no [elemento Resources](resources.md) .

### <a name="description"></a>Descrição

Obrigatório. A descrição da superdica. O **atributo resid** não pode ter mais de 32 caracteres e deve ser definido como o valor do atributo **id** de um elemento **String** no **elemento LongStrings** no [elemento Resources](resources.md) .

> [!NOTE]
> Para Outlook, somente os clientes Windows e Mac suportam o elemento **Description**.

## <a name="example"></a>Exemplo

```xml
<Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
</Supertip>
```

---
title: Elemento Page no arquivo de manifesto
description: O elemento Page define configurações de página HTML que uma função personalizada usa no Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="page-element"></a>Elemento Page

Define as configurações de página HTML usadas por uma função personalizada no Excel.

**Tipo de complemento:** Função Personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md) 

## <a name="attributes"></a>Atributos

Nenhuma

## <a name="child-elements"></a>Elementos filho

|  Elemento  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [SourceLocation](customfunctionssourcelocation.md)  |  Sim  | Cadeia de caracteres com o ID de recurso do arquivo HTML usado por funções personalizadas. |

## <a name="example"></a>Exemplo

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```

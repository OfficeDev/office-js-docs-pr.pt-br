---
title: Elemento SourceLocation para funções personalizadas no arquivo de manifesto
description: Define a localização de um recurso necessário para os elementos de Página ou Script usados por funções personalizadas no Excel.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="sourcelocation-element-custom-functions"></a>Elemento SourceLocation (funções personalizadas)

Define o local de um recurso necessário pelos elementos **Script** ou **Page** usados por funções personalizadas em Excel.

> [!IMPORTANT]
> Este artigo se refere apenas à **SourceLocation** que é filha dos elementos **Page** ou **Script** . Consulte [SourceLocation para](sourcelocation.md) obter informações sobre o **elemento SourceLocation** do manifesto base.

**Tipo de complemento:** Função personalizada

**Válido somente nesses esquemas VersionOverrides**:

- Taskpane 1.0

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [CustomFunctionsRuntime 1.1](../requirement-sets/custom-functions-requirement-sets.md)

## <a name="attributes"></a>Atributos

| Atributo | Obrigatório | Descrição                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| resid     | Sim      | O nome de um recurso de URL definido na seção **Recursos** do manifesto. Não pode ter mais de 32 caracteres. |

## <a name="child-elements"></a>Elementos filho

Nenhum

## <a name="example"></a>Exemplo

```xml
<SourceLocation resid="pageURL"/>
```

import * as ExcelJS from 'exceljs'
import {spawn} from "child_process"
import {create} from "xmlbuilder2"
import {writeFileSync, mkdirSync, existsSync} from 'fs'
import {format, parse, dirname, resolve} from 'path'

function main() {
    const child = spawn("powershell.exe",["./file.ps1"]);
    child.stdout.on("data",function(data){
        const normalizedPath = data.toString().trim()
        excelWork(normalizedPath)
    });
    child.stderr.on("data",function(data){
        console.log("Powershell Errors: " + data);
    });
    child.stdin.end();
}

const TYPES = {
    TEXT: ['TTextValue', 'VARCHAR'],
    STRING: ['TNumberValue', 'VARCHAR'],
    BOOLEAN: ['TCheckValue', 'NUMERIC'],
    DICTIONARY: ['TListValue,TComboValue', 'VARCHAR'],
    DATE: ['TDateValue', 'Timestamptz'],
    FLOAT: ['TNumberValue', 'VARCHAR'],
    INTEGER: ['TNumberValue', 'VARCHAR'],
    LINK: ['TTextValue', 'VARCHAR']
}

async function excelWork(path: string) {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(path)

    const tasksWS = workbook.getWorksheet('Вкладки карточки задачи')
    const summary = workbook.getWorksheet('Summary')
    const dictionary = workbook.getWorksheet('Словари для атрибутов')

    // const finalWorkbook = new ExcelJS.Workbook()
    //
    // if(summary) {
    //     finalWorkbook.addWorksheet('Summary').addRows(summary.getSheetValues())
    // }
    //
    //
    // //Таблица Attributes -> основная ссылка на ВКЗ
    // const attributesSheet = finalWorkbook.addWorksheet('Attributes')

    /**
     * Для группировки значений по имени группы, { 'Гибкие атрибуты задач': ['400001', КОД], ... }
     */

    interface Group {
        entityId: string,
        additionalCode: string,
        keyword: string,
        ancestorKeyword: string
        table: {
            name: string,
            keyFieldName: string
        },
        name: string
        attributeDeclarations: {
            attributeDeclaration:AttributeDeclaration[]
        },
        iucDeclarations: {
            iucDeclaration: IucDeclaration[]
        }
    }

    interface AttributeDeclaration {
        name: string,
        valueType: string,
        fieldName?: string,
        fieldType?: string,
        expression?: string
        flags?: {
            isRequired?: boolean,
            isOverrideStorage?: boolean
        },
        comment?: string
        targetSetDeclarations?: {
            targetSetDeclaration: TargetSetDeclaration[]
        }
    }

    interface TargetSetDeclaration {
        iucKeywords: string
        valueMetadataDfm: string
    }

    interface IucDeclaration {
        keyword: string
    }

    const result: Group[] = []
    summary?.eachRow((row, rowNum) => {
        if(rowNum !== 1) {
            result.push({
                entityId: String(row.getCell('A').value),
                keyword: String(row.getCell('B').value),
                ancestorKeyword: String(row.getCell('D').value),
                table: {
                    name: String(row.getCell('E').value),
                    keyFieldName: String(row.getCell('F').value)
                },
                name: String(row.getCell('G').value),
                attributeDeclarations: {
                    attributeDeclaration: []
                },
                iucDeclarations: {
                    iucDeclaration: []
                },
                //additional data
                additionalCode: String(row.getCell('I').value),
            })
        }
    })
    /**
     * ^ по этим группам генерировать хмлки
     */


    /**
     * Сбор сырых данных без обработки из таблицы ВКЗ, с группировкой по имени группы
     * entityId: data
     */
    const groupedRawData: Record<string, Record<string, any>[]> = {}

    //Основные записи ВКЗ отсортированные по entity
    tasksWS?.eachRow((row, index) => {
        if(index !== 1) {
            const groupName = row.getCell('A').value as string
            const entityId = result.find(group => group.name === groupName)?.entityId
            if (entityId) {
                groupedRawData[entityId] = [...groupedRawData[entityId] ?? [], {
                    // ДЛЯ АТРИБУТОВ
                    name: row.getCell('D').value,
                    value_type: row.getCell('F').value,
                    field_name: row.getCell('D').value,
                    field_type: row.getCell('F').value,
                    is_required: row.getCell('R').value,
                    expression: row.getCell('V').value,
                    comment: row.getCell('W').value,
                    // ДЛЯ TARGETSETS
                    iuc: row.getCell('J').value,
                    caption: row.getCell('E').value,
                    groupIndex: row.getCell('O').value,
                    groupName: row.getCell('N').value,
                    orderIndex: row.getCell('P').value,
                    lineHeight: row.getCell('X').value,
                    readOnlyEdits: row.getCell('Q').value,
                    defaultValue: row.getCell('U').value,
                    multiple: row.getCell('S').value,
                    dictionaryId: row.getCell('G').value,
                }]
            }
        }
    })

    for(const entityId in groupedRawData) {
        const currentGroup = result.find(group => group.entityId === entityId)
        if(currentGroup) {
            let attDeclarations: AttributeDeclaration[] = groupedRawData[entityId]
                // Отсев дублирующихся аттрибутов
                .filter((data, index, arr) =>
                    // конвертирую в массив имен для более простого поиска по значению имени
                    arr.map(item => item.name).indexOf(data.name) === index
                )
                // преобразуем в готовый объект
                .map(rawItem => {
                    const targetSetDeclarations = groupedRawData[entityId]
                        // фильтруем все записи для targetSet по caption
                        .filter(item => item.caption === rawItem.caption)
                        // конвертируем в готовое значение таргет сетов
                        .map(rawSetItem => {
                            const valueMeta = Object.entries(rawSetItem)
                                // отсев всех ненужных ключей и пустых значений
                                .filter(
                                    ([key, value]) => [
                                        'caption',
                                        'groupIndex',
                                        'groupName',
                                        'orderIndex',
                                        'lineHeight',
                                        'readOnlyEdits',
                                        'defaultValue'
                                    ].includes(key) && value)
                                // конкатенация строки по ключу и значению
                                .map(([key, value]) => {
                                    if (key === 'readOnlyEdits') {
                                        return `${key} = ${!Boolean(rawSetItem.readOnlyEdits === 'да') || Boolean(rawSetItem.expression)}`
                                    }
                                    if(['groupIndex', 'orderIndex','lineHeight'].includes(key)) {
                                        // для цыфор
                                        return `${key} = ${value}`
                                    }
                                    // для букав
                                    return `${key} = '${value}'`
                                })

                            if (rawSetItem.field_type === 'DATE') {
                                valueMeta.push('dateTimeKind = dtkDateTime', 'timeKind = tkFullTime')
                            }

                            if(rawSetItem.value_type === 'DICTIONARY') {
                                const values: string[] = []
                                dictionary?.eachRow(row => {
                                    if(row.getCell('A').value === rawItem.dictionaryId) {

                                        values.push(
                                            '  item',
                                            `    Value = '${row.getCell('D')}'`,
                                            `    Desc = '${row.getCell('C')}'`,
                                            '  end'
                                        )
                                    }
                                })

                                valueMeta.push('possibleValues = <', ...values, '>')
                            }

                            return {
                                iucKeywords: rawSetItem.iuc,
                                valueMetadataDfm: valueMeta.reduce((acc, item) => acc.concat('            ', item, '\n'), '\n') + '          '
                            }
                        })



                    const processedItem: AttributeDeclaration = {
                        name: `${currentGroup.additionalCode}${rawItem.name}`,
                        valueType: TYPES[rawItem.value_type as keyof typeof TYPES]?.[0] ?? rawItem.value_type,
                        comment: rawItem.comment,
                        fieldName: `VAL.${rawItem.field_name}`,
                        flags: rawItem.is_required ? {
                                isRequired: rawItem.is_required
                            } : undefined,
                        fieldType: TYPES[rawItem.value_type as keyof typeof TYPES]?.[1] ?? rawItem.value_type,
                        expression: rawItem.expression,
                        targetSetDeclarations:{
                            targetSetDeclaration: targetSetDeclarations
                        }
                    }

                    if(!rawItem.is_required) {
                        delete processedItem.flags
                    }

                    if(rawItem.value_type === 'DICTIONARY') {
                        const [multipleCase, singleCase] = processedItem.valueType.split(',')
                        processedItem.valueType = rawItem.multiple === 'да' ? multipleCase : singleCase
                    }

                    return processedItem
                })

            const defaultValues: AttributeDeclaration[] = [{
                name: 'Id',
                valueType: 'TNumberValue',
                fieldName: 'ID',
                flags: {
                    isOverrideStorage: true
                },
                targetSetDeclarations: {
                    targetSetDeclaration: [{
                        iucKeywords: 'DOCUMENT, NAVIGATOR',
                        valueMetadataDfm: "\n            caption = 'Entity id'\n            visible = false\n          "
                    }]
                },
            }, {
                name: 'Entity',
                valueType: 'TEntityValue',
                flags: {
                    isOverrideStorage: true,
                    isRequired: true
                }
            }, {
                name: 'ObjectName',
                valueType: 'TTextValue',
                comment: 'Entity name',
                fieldName: 'OBJECT_NAME',
                fieldType: 'TTextValue',
                expression: 'Базовая группа гибких атрибутов задач - 40030',
                targetSetDeclarations: {
                    targetSetDeclaration: [{
                        iucKeywords: 'NAVIGATOR',
                        valueMetadataDfm: `caption = 'Entity name'`
                    }]
                }
            }]

            const iucDeclarations = Array.from(new Set(groupedRawData[entityId].map(rawItem => rawItem.iuc)))
                .filter(iuc => !['NAVIGATION', 'DOCUMENT'].includes(iuc))
                .map(uniqIuc => ({keyword: uniqIuc}))



            currentGroup.attributeDeclarations.attributeDeclaration = [...defaultValues, ...attDeclarations]
            currentGroup.iucDeclarations.iucDeclaration = iucDeclarations
            // @ts-ignore
            delete currentGroup.additionalCode
        }
    }


    /**
     * Замена значений с для окончательной таблицы
     */
    // Object.entries(groupedRawData).reduce<Group[]>((acc, group, index) => {
    //     const [entityId, rawValues] = group
    //     rawValues
    //         .filter((value, index, arr) => {
    //             return arr.map(item => item.name).indexOf(value.name) === index
    //         })
    //         .map((item, index) => {
    //             acc.push([
    //                 increment,
    //                 result[groupKey][0],
    //                 `${result[groupKey][1]}${item.name}`,
    //                 TYPES[item.value_type as keyof typeof TYPES]?.[0] ?? item.value_type,
    //                 `VAL.${item.field_name}`,
    //                 TYPES[item.value_type as keyof typeof TYPES]?.[1] ?? item.value_type,
    //                 item.is_required === 'да' ? 'true' : null,
    //                 item.expression,
    //                 item.caption
    //             ])
    //             increment += 1
    //         })
    //     increment = (index + 2) * 150
    //     return acc
    //
    // }, [])
    //
    // // TargetSets
    // // const targetSetsSheet = finalWorkbook.addWorksheet('TargetSets')
    //
    // const targetSetsData = Object.values(groupedAttributesData).reduce<any[][]>((acc, groupValues) => {
    //     groupValues.map(item => {
    //         const id = processedAttributesData.find(itemToFind => itemToFind?.[8] === item.caption)?.[0]
    //         acc.push([
    //             id,
    //             item.iuc,
    //             item.caption,
    //             item.groupIndex,
    //             item.groupName,
    //             item.orderIndex,
    //             undefined,
    //             (!Boolean(item.readOnlyEdits === 'да') || Boolean(item.expression)).toString(),
    //             undefined,
    //             undefined,
    //             item.field_type === 'DATE' ? 'dtkDate' : undefined,
    //             item.field_type === 'DATE' ? 'tkFullTime' : undefined,
    //             item.defaultValue
    //         ])
    //     })
    //     return acc
    // }, [])

    // const attTable = attributesSheet.addTable({
    //     name: 'Attributes',
    //     ref: 'A1',
    //     headerRow: true,
    //     style: {
    //         theme: 'TableStyleMedium2',
    //         showRowStripes: true,
    //     },
    //     columns: [
    //         { name: 'id', filterButton: true},
    //         { name: 'group_id', filterButton: true },
    //         { name: 'name', filterButton: true },
    //         { name: 'value_type', filterButton: true },
    //         { name: 'field_name', filterButton: true },
    //         { name: 'field_type', filterButton: true },
    //         { name: 'is_required', filterButton: true },
    //         { name: 'expression', filterButton: true },
    //         { name: 'caption' }
    //     ],
    //     rows: processedAttributesData
    // })
    //
    // targetSetsSheet.addTable({
    //     name: 'TargetSets',
    //     ref: 'A1',
    //     headerRow: true,
    //     style: {
    //         theme: 'TableStyleMedium2',
    //         showRowStripes: true,
    //     },
    //     columns: [
    //         { name: 'attr_id', filterButton: true},
    //         { name: 'iuc', filterButton: true },
    //         { name: 'caption', filterButton: true },
    //         { name: 'groupIndex', filterButton: true },
    //         { name: 'groupName', filterButton: true },
    //         { name: 'orderIndex', filterButton: true },
    //         { name: 'linesHeight', filterButton: true },
    //         { name: 'readOnlyEdits', filterButton: true },
    //         { name: 'precision', filterButton: true },
    //         { name: 'step', filterButton: true },
    //         { name: 'dateTimeKind', filterButton: true },
    //         { name: 'timeKind', filterButton: true },
    //         { name: 'initString', filterButton: true },
    //     ],
    //     rows: targetSetsData
    // })
    //
    // attTable.removeColumns(8, 1)
    //
    // attTable.commit()
    // await finalWorkbook.xlsx.writeFile(finalPath);

    const template = {
        '@xsi:schemaLocation': 'http://argustelecom.ru/inventory/model-metadata http://argustelecom.ru/inventory/model-metadata/metamodel_1_0.xsd',
            '@xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
            '@xmlns:ns': 'http://argustelecom.ru/inventory/model-metadata'
    }

    const dirPath = resolve(dirname(path), 'models')
    if(!existsSync(dirPath)) {
        mkdirSync(dirPath)
    }

    result.forEach(builderRes => {
        const builder = create({'ns:entity': {...template, ...builderRes}}).end({prettyPrint: true, format: 'xml'})
        const keyword = builderRes.keyword.trim() === '' ? 'UnknownKeyword' : builderRes.keyword.trim()
        const parsedPath = parse(path)
        const finalPath = format({...parsedPath, base: `${keyword}.xml`, dir: parsedPath.dir + '\\models'})

        writeFileSync(finalPath, builder)
    })

}

main()

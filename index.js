"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const ExcelJS = __importStar(require("exceljs"));
const child_process_1 = require("child_process");
const xmlbuilder2_1 = require("xmlbuilder2");
const fs_1 = require("fs");
const path_1 = require("path");
function main() {
    const child = (0, child_process_1.spawn)("powershell.exe", ["./file.ps1"]);
    child.stdout.on("data", function (data) {
        const normalizedPath = data.toString().trim();
        excelWork(normalizedPath);
    });
    child.stderr.on("data", function (data) {
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
};
async function excelWork(path) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path);
    const tasksWS = workbook.getWorksheet('Вкладки карточки задачи');
    const summary = workbook.getWorksheet('Summary');
    const dictionary = workbook.getWorksheet('Словари для атрибутов');
    const result = [];
    summary?.eachRow((row, rowNum) => {
        if (rowNum !== 1) {
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
            });
        }
    });
    /**
     * ^ по этим группам генерировать хмлки
     */
    /**
     * Сбор сырых данных без обработки из таблицы ВКЗ, с группировкой по имени группы
     * entityId: data
     */
    const groupedRawData = {};
    //Основные записи ВКЗ отсортированные по entity
    tasksWS?.eachRow((row, index) => {
        if (index !== 1) {
            const groupName = row.getCell('A').value;
            const entityId = result.find(group => group.name === groupName)?.entityId;
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
                    }];
            }
        }
    });
    for (const entityId in groupedRawData) {
        const currentGroup = result.find(group => group.entityId === entityId);
        if (currentGroup) {
            let attDeclarations = groupedRawData[entityId]
                // Отсев дублирующихся аттрибутов
                .filter((data, index, arr) => 
            // конвертирую в массив имен для более простого поиска по значению имени
            arr.map(item => item.name).indexOf(data.name) === index)
                // преобразуем в готовый объект
                .map(rawItem => {
                const targetSetDeclarations = groupedRawData[entityId]
                    // фильтруем все записи для targetSet по caption
                    .filter(item => item.field_name === rawItem.field_name)
                    // конвертируем в готовое значение таргет сетов
                    .map(rawSetItem => {
                    const valueMeta = Object.entries(rawSetItem)
                        // отсев всех ненужных ключей и пустых значений
                        .filter(([key, value]) => [
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
                            return `${key} = ${!Boolean(rawSetItem.readOnlyEdits === 'да') || Boolean(rawSetItem.expression)}`;
                        }
                        if (['groupIndex', 'orderIndex', 'lineHeight'].includes(key)) {
                            // для цыфор
                            return `${key} = ${value}`;
                        }
                        // для букав
                        return `${key} = '${value}'`;
                    });
                    if (rawSetItem.field_type === 'DATE') {
                        valueMeta.push('dateTimeKind = dtkDateTime', 'timeKind = tkFullTime');
                    }
                    if (rawSetItem.value_type === 'DICTIONARY') {
                        const values = [];
                        dictionary?.eachRow(row => {
                            if (row.getCell('A').value === rawItem.dictionaryId) {
                                values.push('  item', `    Value = '${row.getCell('D')}'`, `    Desc = '${row.getCell('C')}'`, '  end');
                            }
                        });
                        valueMeta.push('possibleValues = <', ...values, '>');
                    }
                    return {
                        iucKeywords: rawSetItem.iuc,
                        valueMetadataDfm: valueMeta.reduce((acc, item) => acc.concat('            ', item, '\n'), '\n') + '          '
                    };
                });
                const processedItem = {
                    name: `${currentGroup.additionalCode}${rawItem.name}`,
                    valueType: TYPES[rawItem.value_type]?.[0] ?? rawItem.value_type,
                    comment: rawItem.comment,
                    fieldName: `VAL.${rawItem.field_name}`,
                    fieldType: TYPES[rawItem.value_type]?.[1] ?? rawItem.value_type,
                    expression: rawItem.expression,
                    flags: {
                        isRequired: rawItem.is_required === 'да'
                    },
                    targetSetDeclarations: {
                        targetSetDeclaration: targetSetDeclarations
                    }
                };
                if (!processedItem.flags?.isRequired) {
                    delete processedItem.flags;
                }
                if (rawItem.value_type === 'DICTIONARY') {
                    const [multipleCase, singleCase] = processedItem.valueType.split(',');
                    processedItem.valueType = rawItem.multiple === 'да' ? multipleCase : singleCase;
                }
                return processedItem;
            });
            const iucDeclarations = Array.from(new Set(groupedRawData[entityId].map(rawItem => rawItem.iuc)))
                .filter(iuc => !['NAVIGATION', 'DOCUMENT'].includes(iuc))
                .map(uniqIuc => ({ keyword: uniqIuc }));
            const defaultValues = [{
                    name: 'Id',
                    valueType: 'TNumberValue',
                    fieldName: 'ID',
                    flags: {
                        isOverrideStorage: true
                    },
                    targetSetDeclarations: {
                        targetSetDeclaration: [{
                                iucKeywords: ['DOCUMENT', 'NAVIGATOR', ...iucDeclarations].join(', '),
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
                }];
            currentGroup.attributeDeclarations.attributeDeclaration = [...defaultValues, ...attDeclarations];
            currentGroup.iucDeclarations.iucDeclaration = iucDeclarations;
            // @ts-ignore
            delete currentGroup.additionalCode;
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
    };
    const dirPath = (0, path_1.resolve)((0, path_1.dirname)(path), 'models');
    if (!(0, fs_1.existsSync)(dirPath)) {
        (0, fs_1.mkdirSync)(dirPath);
    }
    result.forEach(builderRes => {
        const builder = (0, xmlbuilder2_1.create)({ 'ns:entity': { ...template, ...builderRes } }).end({ prettyPrint: true, format: 'xml' });
        const keyword = builderRes.entityId;
        const parsedPath = (0, path_1.parse)(path);
        const finalPath = (0, path_1.format)({ ...parsedPath, base: `${keyword}.xml`, dir: parsedPath.dir + '\\models' });
        (0, fs_1.writeFileSync)(finalPath, builder);
    });
}
main();

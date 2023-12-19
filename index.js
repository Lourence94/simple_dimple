"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
Object.defineProperty(exports, "__esModule", { value: true });
var ExcelJS = require("exceljs");
var child_process_1 = require("child_process");
var xmlbuilder2_1 = require("xmlbuilder2");
var fs_1 = require("fs");
var path_1 = require("path");
function main() {
    var child = (0, child_process_1.spawn)("powershell.exe", ["./file.ps1"]);
    child.stdout.on("data", function (data) {
        var normalizedPath = data.toString().trim();
        excelWork(normalizedPath);
    });
    child.stderr.on("data", function (data) {
        console.log("Powershell Errors: " + data);
    });
    child.stdin.end();
}
var TYPES = {
    TEXT: ['TTextValue', 'VARCHAR'],
    STRING: ['TNumberValue', 'VARCHAR'],
    BOOLEAN: ['TCheckValue', 'NUMERIC'],
    DICTIONARY: ['TListValue,TComboValue', 'VARCHAR'],
    DATE: ['TDateValue', 'Timestamptz'],
    FLOAT: ['TNumberValue', 'VARCHAR'],
    INTEGER: ['TNumberValue', 'VARCHAR'],
    LINK: ['TTextValue', 'VARCHAR']
};
function excelWork(path) {
    return __awaiter(this, void 0, void 0, function () {
        var workbook, tasksWS, summary, dictionary, result, groupedRawData, _loop_1, entityId, template, dirPath;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    workbook = new ExcelJS.Workbook();
                    return [4 /*yield*/, workbook.xlsx.readFile(path)];
                case 1:
                    _a.sent();
                    tasksWS = workbook.getWorksheet('Вкладки карточки задачи');
                    summary = workbook.getWorksheet('Summary');
                    dictionary = workbook.getWorksheet('Словари для атрибутов');
                    result = [];
                    summary === null || summary === void 0 ? void 0 : summary.eachRow(function (row, rowNum) {
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
                    groupedRawData = {};
                    //Основные записи ВКЗ отсортированные по entity
                    tasksWS === null || tasksWS === void 0 ? void 0 : tasksWS.eachRow(function (row, index) {
                        var _a, _b;
                        if (index !== 1) {
                            var groupName_1 = row.getCell('A').value;
                            var entityId = (_a = result.find(function (group) { return group.name === groupName_1; })) === null || _a === void 0 ? void 0 : _a.entityId;
                            if (entityId) {
                                groupedRawData[entityId] = __spreadArray(__spreadArray([], (_b = groupedRawData[entityId]) !== null && _b !== void 0 ? _b : [], true), [{
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
                                    }], false);
                            }
                        }
                    });
                    _loop_1 = function (entityId) {
                        var currentGroup = result.find(function (group) { return group.entityId === entityId; });
                        if (currentGroup) {
                            var attDeclarations = groupedRawData[entityId]
                                // Отсев дублирующихся аттрибутов
                                .filter(function (data, index, arr) {
                                // конвертирую в массив имен для более простого поиска по значению имени
                                return arr.map(function (item) { return item.name; }).indexOf(data.name) === index;
                            })
                                // преобразуем в готовый объект
                                .map(function (rawItem) {
                                var _a, _b, _c, _d;
                                var targetSetDeclarations = groupedRawData[entityId]
                                    // фильтруем все записи для targetSet по caption
                                    .filter(function (item) { return item.field_name === rawItem.field_name; })
                                    // конвертируем в готовое значение таргет сетов
                                    .map(function (rawSetItem) {
                                    var valueMeta = Object.entries(rawSetItem)
                                        // отсев всех ненужных ключей и пустых значений
                                        .filter(function (_a) {
                                        var key = _a[0], value = _a[1];
                                        return [
                                            'caption',
                                            'groupIndex',
                                            'groupName',
                                            'orderIndex',
                                            'lineHeight',
                                            'readOnlyEdits',
                                            'defaultValue'
                                        ].includes(key) && value;
                                    })
                                        // конкатенация строки по ключу и значению
                                        .map(function (_a) {
                                        var key = _a[0], value = _a[1];
                                        if (key === 'readOnlyEdits') {
                                            return "".concat(key, " = ").concat(!Boolean(rawSetItem.readOnlyEdits === 'да') || Boolean(rawSetItem.expression));
                                        }
                                        if (['groupIndex', 'orderIndex', 'lineHeight'].includes(key)) {
                                            // для цыфор
                                            return "".concat(key, " = ").concat(value);
                                        }
                                        // для букав
                                        return "".concat(key, " = '").concat(value, "'");
                                    });
                                    if (rawSetItem.field_type === 'DATE') {
                                        valueMeta.push('dateTimeKind = dtkDateTime', 'timeKind = tkFullTime');
                                    }
                                    if (rawSetItem.value_type === 'DICTIONARY') {
                                        var values_1 = [];
                                        dictionary === null || dictionary === void 0 ? void 0 : dictionary.eachRow(function (row) {
                                            if (row.getCell('A').value === rawItem.dictionaryId) {
                                                values_1.push('  item', "    Value = '".concat(row.getCell('D'), "'"), "    Desc = '".concat(row.getCell('C'), "'"), '  end');
                                            }
                                        });
                                        valueMeta.push.apply(valueMeta, __spreadArray(__spreadArray(['possibleValues = <'], values_1, false), ['>'], false));
                                    }
                                    return {
                                        iucKeywords: rawSetItem.iuc,
                                        valueMetadataDfm: valueMeta.reduce(function (acc, item) { return acc.concat('            ', item, '\n'); }, '\n') + '          '
                                    };
                                });
                                var processedItem = {
                                    name: "".concat(currentGroup.additionalCode).concat(rawItem.name),
                                    valueType: (_b = (_a = TYPES[rawItem.value_type]) === null || _a === void 0 ? void 0 : _a[0]) !== null && _b !== void 0 ? _b : rawItem.value_type,
                                    comment: rawItem.comment,
                                    fieldName: "VAL.".concat(rawItem.field_name),
                                    flags: rawItem.is_required === 'да' ? {
                                        isRequired: rawItem.is_required === 'да'
                                    } : undefined,
                                    fieldType: (_d = (_c = TYPES[rawItem.value_type]) === null || _c === void 0 ? void 0 : _c[1]) !== null && _d !== void 0 ? _d : rawItem.value_type,
                                    expression: rawItem.expression,
                                    targetSetDeclarations: {
                                        targetSetDeclaration: targetSetDeclarations
                                    }
                                };
                                if (!Boolean(rawItem.is_required === 'да')) {
                                    delete processedItem.flags;
                                }
                                if (rawItem.value_type === 'DICTIONARY') {
                                    var _e = processedItem.valueType.split(','), multipleCase = _e[0], singleCase = _e[1];
                                    processedItem.valueType = rawItem.multiple === 'да' ? multipleCase : singleCase;
                                }
                                return processedItem;
                            });
                            var defaultValues = [{
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
                                                valueMetadataDfm: "caption = 'Entity name'"
                                            }]
                                    }
                                }];
                            var iucDeclarations = Array.from(new Set(groupedRawData[entityId].map(function (rawItem) { return rawItem.iuc; })))
                                .filter(function (iuc) { return !['NAVIGATION', 'DOCUMENT'].includes(iuc); })
                                .map(function (uniqIuc) { return ({ keyword: uniqIuc }); });
                            currentGroup.attributeDeclarations.attributeDeclaration = __spreadArray(__spreadArray([], defaultValues, true), attDeclarations, true);
                            currentGroup.iucDeclarations.iucDeclaration = iucDeclarations;
                            // @ts-ignore
                            delete currentGroup.additionalCode;
                        }
                    };
                    for (entityId in groupedRawData) {
                        _loop_1(entityId);
                    }
                    template = {
                        '@xsi:schemaLocation': 'http://argustelecom.ru/inventory/model-metadata http://argustelecom.ru/inventory/model-metadata/metamodel_1_0.xsd',
                        '@xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
                        '@xmlns:ns': 'http://argustelecom.ru/inventory/model-metadata'
                    };
                    dirPath = (0, path_1.resolve)((0, path_1.dirname)(path), 'models');
                    if (!(0, fs_1.existsSync)(dirPath)) {
                        (0, fs_1.mkdirSync)(dirPath);
                    }
                    result.forEach(function (builderRes) {
                        var builder = (0, xmlbuilder2_1.create)({ 'ns:entity': __assign(__assign({}, template), builderRes) }).end({ prettyPrint: true, format: 'xml' });
                        var keyword = builderRes.entityId;
                        var parsedPath = (0, path_1.parse)(path);
                        var finalPath = (0, path_1.format)(__assign(__assign({}, parsedPath), { base: "".concat(keyword, ".xml"), dir: parsedPath.dir + '\\models' }));
                        (0, fs_1.writeFileSync)(finalPath, builder);
                    });
                    return [2 /*return*/];
            }
        });
    });
}
main();

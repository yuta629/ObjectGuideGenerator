/*******************************************************************************
 *
 * InDesign オブジェクトガイド生成スクリプト
 *
 * 概要:
 * 選択されたオブジェクトを基準にページガイドを生成します。
 * Illustratorの「ガイドを作成」機能に似た動作を目指します。
 *
 * バージョン: 2.4.0 (下位互換性 安定版)
 * 作成日: 2024/06/07
 * 更新日: 2025/06/09
 * 修正内容:
 * - 古いExtendScript環境にない`Array.prototype.indexOf`のPolyfillを追加。
 * - 古いExtendScript環境にない`Object.keys`のPolyfillを追加。
 * - 「グループとして」「オブジェクトごと」両方の処理経路で重複回避が機能するよう修正。
 * 対応バージョン: Adobe InDesign (CS6以降での動作を想定)
 *
 ******************************************************************************/

(function() {

    // Object.keysのPolyfill (古いInDesignバージョン対応)
    if (!Object.keys) {
        Object.keys = (function() {
            'use strict';
            var hasOwnProperty = Object.prototype.hasOwnProperty,
                hasDontEnumBug = !({ toString: null }).propertyIsEnumerable('toString'),
                dontEnums = [
                    'toString', 'toLocaleString', 'valueOf', 'hasOwnProperty',
                    'isPrototypeOf', 'propertyIsEnumerable', 'constructor'
                ],
                dontEnumsLength = dontEnums.length;

            return function(obj) {
                if (typeof obj !== 'object' && (typeof obj !== 'function' || obj === null)) {
                    throw new TypeError('Object.keys called on non-object');
                }
                var result = [], prop, i;
                for (prop in obj) {
                    if (hasOwnProperty.call(obj, prop)) {
                        result.push(prop);
                    }
                }
                if (hasDontEnumBug) {
                    for (i = 0; i < dontEnumsLength; i++) {
                        if (hasOwnProperty.call(obj, dontEnums[i])) {
                            result.push(dontEnums[i]);
                        }
                    }
                }
                return result;
            };
        }());
    }

    // Array.prototype.indexOf の Polyfill (古いInDesignバージョン対応)
    if (!Array.prototype.indexOf) {
        Array.prototype.indexOf = function(searchElement, fromIndex) {
            var k;
            if (this == null) {
                throw new TypeError('"this" is null or not defined');
            }
            var o = Object(this);
            var len = o.length >>> 0;
            if (len === 0) {
                return -1;
            }
            var n = fromIndex | 0;
            if (n >= len) {
                return -1;
            }
            k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);
            while (k < len) {
                if (k in o && o[k] === searchElement) {
                    return k;
                }
                k++;
            }
            return -1;
        };
    }


    // =============================================================================
    // メイン処理
    // =============================================================================
    function main() {
        if (app.documents.length === 0) {
            alert("処理対象のドキュメントを開いてください。");
            return;
        }
        if (app.selection.length === 0) {
            alert("ガイドを生成するオブジェクトを選択してください。");
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = app.selection;

        var userSettings = showDialog();
        if (userSettings === null) {
            return;
        }

        app.doScript(
            function() {
                run(doc, originalSelection, userSettings);
            },
            ScriptLanguage.JAVASCRIPT,
            [],
            UndoModes.ENTIRE_SCRIPT,
            "オブジェクトからガイド生成"
        );
    }

    /**
     * スクリプトの主処理を実行する
     */
    function run(doc, selection, settings) {
        var objectsByPage = {}; 

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            var page = isItemOnPage(item);
            if (page !== null) {
                if (!objectsByPage[page.id]) {
                    objectsByPage[page.id] = [];
                }
                objectsByPage[page.id].push(item);
            }
        }
        
        if (Object.keys(objectsByPage).length === 0) {
            alert("ページ内に完全に配置されたオブジェクトを選択してください。");
            return;
        }

        var guideCount = 0;
        var targetLayer = getTargetLayer(doc, settings.createLayer);
        
        var hLocations = []; // 水平ガイドの座標記録用
        var vLocations = []; // 垂直ガイドの座標記録用

        if (settings.processAsGroup) {
            for (var pageId in objectsByPage) {
                if (objectsByPage.hasOwnProperty(pageId)) {
                    var itemsOnPage = objectsByPage[pageId];
                    var targetPage = itemsOnPage[0].parentPage; 
                    guideCount += createGuidesForItemGroup(itemsOnPage, settings, targetPage, targetLayer, hLocations, vLocations);
                }
            }
        } else {
            for (var pageId in objectsByPage) {
                 if (objectsByPage.hasOwnProperty(pageId)) {
                    var itemsOnPage = objectsByPage[pageId];
                    for (var j = 0; j < itemsOnPage.length; j++){
                        guideCount += createGuidesForSingleItem(itemsOnPage[j], settings, targetLayer, hLocations, vLocations);
                    }
                }
            }
        }

        if (guideCount > 0) {
            alert(guideCount + "本のガイドを生成しました。");
        } else {
            alert("指定された条件で生成できるガイドはありませんでした。");
        }
    }


    // =============================================================================
    // ダイアログ関連
    // =============================================================================
    function showDialog() {
        var dialog = new Window("dialog", "オブジェクトからガイド生成");
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        var processPanel = dialog.add("panel", undefined, "処理対象");
        processPanel.orientation = "column";
        processPanel.alignChildren = ["left", "top"];
        var rbGroup = processPanel.add("radiobutton", undefined, "グループとして");
        var rbIndividual = processPanel.add("radiobutton", undefined, "オブジェクトごと");
        rbIndividual.value = true;

        var positionPanel = dialog.add("panel", undefined, "生成するガイド");
        positionPanel.orientation = "row";
        var posGroup1 = positionPanel.add("group");
        posGroup1.orientation = "column";
        posGroup1.alignChildren = ["left", "top"];
        var cbTop = posGroup1.add("checkbox", undefined, "上辺");
        cbTop.value = true;
        var cbBottom = posGroup1.add("checkbox", undefined, "下辺");
        cbBottom.value = true;
        var cbHCenter = posGroup1.add("checkbox", undefined, "水平センター");

        var posGroup2 = positionPanel.add("group");
        posGroup2.orientation = "column";
        posGroup2.alignChildren = ["left", "top"];
        var cbLeft = posGroup2.add("checkbox", undefined, "左辺");
        cbLeft.value = true;
        var cbRight = posGroup2.add("checkbox", undefined, "右辺");
        cbRight.value = true;
        var cbVCenter = posGroup2.add("checkbox", undefined, "垂直センター");

        var optionPanel = dialog.add("panel", undefined, "オプション");
        optionPanel.orientation = "column";
        optionPanel.alignChildren = ["left", "top"];
        var cbCreateLayer = optionPanel.add("checkbox", undefined, "専用レイヤーを作成（'生成ガイド'）");
        cbCreateLayer.value = true;

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "right";
        buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
        buttonGroup.add("button", undefined, "OK", { name: "ok" });

        if (dialog.show() == 1) { // OK
            return {
                processAsGroup: rbGroup.value,
                top: cbTop.value,
                bottom: cbBottom.value,
                hCenter: cbHCenter.value,
                left: cbLeft.value,
                right: cbRight.value,
                vCenter: cbVCenter.value,
                createLayer: cbCreateLayer.value
            };
        }
        return null; // Cancel
    }

    // =============================================================================
    // ガイド生成ロジック
    // =============================================================================
    function createGuidesForSingleItem(item, settings, layer, hLocations, vLocations) {
        var page = isItemOnPage(item);
        if (page) {
            var bounds = item.visibleBounds;
            return createGuidesFromBounds(bounds, settings, page, layer, hLocations, vLocations);
        }
        return 0;
    }

    function createGuidesForItemGroup(items, settings, page, layer, hLocations, vLocations) {
        if (items.length === 0) return 0;
        
        var firstBounds = items[0].visibleBounds;
        var y1 = firstBounds[0], x1 = firstBounds[1], y2 = firstBounds[2], x2 = firstBounds[3];
        for (var i = 1; i < items.length; i++) {
            var bounds = items[i].visibleBounds;
            y1 = Math.min(y1, bounds[0]);
            x1 = Math.min(x1, bounds[1]);
            y2 = Math.max(y2, bounds[2]);
            x2 = Math.max(x2, bounds[3]);
        }
        var combinedBounds = [y1, x1, y2, x2];

        return createGuidesFromBounds(combinedBounds, settings, page, layer, hLocations, vLocations);
    }

    function createGuidesFromBounds(bounds, settings, page, layer, hLocations, vLocations) {
        var y1 = bounds[0], x1 = bounds[1], y2 = bounds[2], x2 = bounds[3];
        var hCenter = (y1 + y2) / 2;
        var vCenter = (x1 + x2) / 2;
        var count = 0;

        if (settings.top && hLocations.indexOf(y1) === -1) { createGuide(page, layer, "horizontal", y1); hLocations.push(y1); count++; }
        if (settings.bottom && hLocations.indexOf(y2) === -1) { createGuide(page, layer, "horizontal", y2); hLocations.push(y2); count++; }
        if (settings.hCenter && hLocations.indexOf(hCenter) === -1) { createGuide(page, layer, "horizontal", hCenter); hLocations.push(hCenter); count++; }
        if (settings.left && vLocations.indexOf(x1) === -1) { createGuide(page, layer, "vertical", x1); vLocations.push(x1); count++; }
        if (settings.right && vLocations.indexOf(x2) === -1) { createGuide(page, layer, "vertical", x2); vLocations.push(x2); count++; }
        if (settings.vCenter && vLocations.indexOf(vCenter) === -1) { createGuide(page, layer, "vertical", vCenter); vLocations.push(vCenter); count++; }
        
        return count;
    }


    // =============================================================================
    // ヘルパー関数
    // =============================================================================
    function isItemOnPage(item) {
        try {
            if (item.parentPage !== null && item.parentPage.isValid) {
                return item.parentPage;
            }
        } catch (e) {
            return null;
        }
        return null;
    }

    function getTargetLayer(doc, shouldCreate) {
        if (!shouldCreate) {
            return doc.activeLayer;
        }
        var layerName = "生成ガイド";
        var layer = doc.layers.itemByName(layerName);
        if (layer.isValid) {
            return layer;
        }
        return doc.layers.add({ name: layerName });
    }

    function createGuide(page, layer, orientation, location) {
        page.guides.add(layer, {
            orientation: (orientation === "horizontal") ? HorizontalOrVertical.HORIZONTAL : HorizontalOrVertical.VERTICAL,
            location: location
        });
    }

    main();

})();
import * as Tools from "./Tools.js";
import * as Enums from "./Enums.js";
import * as Menu from "./Menu.js";
import * as PropertyGrid from "./PropertyGrid.js";
import * as Window from "./Window.js";
import * as Terminus from "./Terminus.js";
import * as Search from "./Search.js";
import * as GlobalEvents from "./GlobalEvents.js";

const tooltipOffset = 15;

window.dotnet = chrome.webview.hostObjects.dotnet;
window.syncDotnet = chrome.webview.hostObjects.sync.dotnet;
window.tooltip = document.getElementById("app-tooltip");
window.tooltipId = 0;
window.updateInstances = updateInstances;
window.editAddress = editAddress;
window.renameEntry = renameEntry;
window.applyTheme = applyTheme;
window.navigate = navigate;
window.refreshEntries = refreshEntries;
window.showWindow = showWindow;
window.showPropertyGrid = showPropertyGrid;
window.showTerminal = showTerminal;
window.closeTerminal = closeTerminal;
window.setThumbnailsSize = setThumbnailsSize;
window.serverUrl = syncDotnet.serverUrl;
window.getEntriesRect = getEntriesRect;
window.showAlert = showAlert;
window.showToast = showToast;
window.goBack = goBack;
window.goForward = goForward;
window.selectParsingNameEntry = selectParsingNameEntry;
window.getParsingNameAndSelectionAtPoint = getParsingNameAndSelectionAtPoint;
window.getSelection = getSelection;
window.openSearch = openSearch;
window.openMenu = openMenu;
window.windowHandle = syncDotnet.getWindowHandle();

document.getElementById("app-filter-type").innerHTML = Tools.Resource("ViewFilter");
const viewFilterInput = document.getElementById("app-filter-value");
const filterDiv = document.getElementById("app-filter");
const entriesParent = document.getElementById("app-entries-parent");
const appThumbnailsSize = document.getElementById("app-thumbnails-size");
appThumbnailsSize.onmousemove = () => showAppThumnbailSizeMenu();

// view settings
const instanceSettings = JSON.parse(syncDotnet.getInstanceSettings());
window.instanceName = instanceSettings.instanceName;
window.paging = instanceSettings.paging;
window.entryEnumerateOptions = instanceSettings.entryEnumerateOptions;
window.viewBy = instanceSettings.viewBy;
window.viewByImageOptions = instanceSettings.viewByImageOptions;
window.viewByDetailsOptions = instanceSettings.viewByDetailsOptions;
window.sortBy = instanceSettings.sortBy;
window.sortDirection = instanceSettings.sortDirection;
window.showDevTools = instanceSettings.showDevTools;
window.isAdministrator = syncDotnet.isAdministrator;
saveThumbnailsSize(instanceSettings.thumbnailsSize);

// selection
let selection = {};
let selectionIndex = -1;

let totalEntries = 0;

// edition
let isEditing = false;

GlobalEvents.GlobalEvents.install();

window.GlobalEvents.addEventListener("ContextMenuNavigation", (e) => {
    const shift = e.detail.srcEvent.shiftKey;
    const id = parseInt(e.detail.id);
    if (id !== 0 && !id)
        return;

    // rotate menus
    let newId = shift ? id - 1 : id + 1;
    if (newId === Enums.MenuId.AppThumbnailSize && window.viewBy != Enums.ViewBy.Images) { // special images cases
        newId = shift ? newId - 1 : newId + 1;
    }

    if (newId < 0) {
        newId = Enums.MenuId.MaxId - 1;
    }
    else if (newId >= Enums.MenuId.MaxId) {
        newId = 0;
    }

    openMenu(newId);
});

updateThumbnailsSizeControls();
updateInstances();

const oldLog = console.log;
console.log = function () {
    const args = Array.prototype.slice.call(arguments);
    dotnet.sendEvent(Enums.WebEventType.ConsoleLog, { arguments: args });
    oldLog.apply(console, arguments);
};

window.onerror = (message, source, lineno, colno, error) => { dotnet.onError(message, source, lineno, colno, error?.stack); }

document.addEventListener("readystatechange", () => {
    dotnet.sendEvent(Enums.WebEventType.ReadyStateChange, { state: document.readyState, message: `ShellBat JS initialized on navigator: ${navigator.userAgent}` });
});

if (window.isAdministrator) {
    const appAdmin = document.getElementById("app-administrator");
    appAdmin.innerText = Tools.Resource("AdministratorMode");
    appAdmin.style.display = "block";
}

document.getElementById("app-filter-close").onclick = e => {
    clearViewFilter();
}

window.onresize = resize;
window.visualViewport.addEventListener("resize", (e) => {
    Menu.Menu.dismissRootMenu();
    dotnet.sendEvent(Enums.WebEventType.VisualViewportChanged, { scale: window.devicePixelRatio });
    if (window.devicePixelRatio == 1) {
        document.getElementById("app-zoom").innerText = "";
    }
    else {
        document.getElementById("app-zoom").innerText = Math.round(window.devicePixelRatio * 100) + "%";
    }
});

window.addEventListener("drop", (e) => {
    e.preventDefault();
})

window.addEventListener("dragover", (e) => {
    e.preventDefault();
})

entriesParent.addEventListener("contextmenu", (e) => {
    Menu.Menu.dismissRootMenu();
    entryContextMenu(e);
    e.preventDefault();
});

entriesParent.addEventListener("scroll", (e) => {
    highlightViewFilterText();
    Menu.Menu.dismissRootMenu();
});

viewFilterInput.addEventListener("input", (e) => {
    Menu.Menu.dismissRootMenu();
    filterEntries();
    highlightViewFilterText();
});

viewFilterInput.addEventListener("focusout", (e) => {
    if (viewFilterInput.value.trim() == "") {
        clearViewFilter();
    }
});

// webview2 eats key events, so we need to forward them manually
document.addEventListener("keydown", (e) => {
    if (isEditing)
        return;

    const ae = document.activeElement;
    if (ae == viewFilterInput) {
        if (e.code == "Escape") {
            clearViewFilter();
            e.preventDefault();
        }

        // let these pass
        if (e.code != "ArrowDown" && e.code != "ArrowUp")
            return;
    }
    else if (ae.tagName == "INPUT" || ae.tagName == "TEXTAREA" || ae.isContentEditable) {
        return;
    }

    if (Menu.Menu.handleKeydown(e)) {
        e.preventDefault();
        return;
    }

    if (moveSelection(e)) {
        e.preventDefault();
        return;
    }

    if (e.code == "Enter") {
        entryDoubleClicked(e);
        e.preventDefault();
        return;
    }

    if (e.code == "Backspace") {
        goBack();
        e.preventDefault();
        return;
    }

    if (e.code == "ContextMenu" ||
        e.code == "F10") {
        entryContextMenu(e);
        e.preventDefault();
        return;
    }

    if (e.code == "Delete") {
        deleteSelectedEntries();
        e.preventDefault();
        return;
    }

    if (e.code == "F2") {
        renameEntry();
        e.preventDefault();
        return;
    }

    if (e.code == "Escape") {
        clearSelection();
        clearViewFilter();
        e.preventDefault();
        return;
    }

    if (updateViewFilter(e)) {
        e.preventDefault();
        return;
    }

    dotnet.sendEvent(Enums.WebEventType.KeyDown, { code: e.code, key: e.key, type: e.type, shift: e.shiftKey, ctrl: e.ctrlKey, alt: e.altKey, meta: e.metaKey });
});

document.addEventListener("keyup", (e) => {
    dotnet.sendEvent(Enums.WebEventType.KeyUp, { code: e.code, key: e.key, type: e.type, shift: e.shiftKey, ctrl: e.ctrlKey, alt: e.altKey, meta: e.metaKey });
});

document.getElementById("app-close-button").onclick = () => dotnet.sendEvent(Enums.WebEventType.Close);
document.getElementById("app-max-button").onclick = () => dotnet.sendEvent(Enums.WebEventType.MaximizeRestore);
document.getElementById("app-min-button").onclick = () => dotnet.sendEvent(Enums.WebEventType.Minimize);
document.getElementById("app-recycle-bin").onclick = () => navigate("::{645FF040-5081-101B-9F08-00AA002F954E}");
document.getElementById("app-open-with-explorer").onclick = () => dotnet.openWithExplorer(getSelection());

const appInfo = document.getElementById("app-info-button");
appInfo.onmousemove = () => showAppInfoMenu();

const appHistory = document.getElementById("app-history");
appHistory.onmousemove = () => showHistory();

const appFavorites = document.getElementById("app-favorites");
appFavorites.onmousemove = () => showFavorites();

document.getElementById("app-go-back").onclick = () => goBack();
document.getElementById("app-go-forward").onclick = () => goForward();
document.getElementById("app-go-up").onclick = () => navigate(window.upParsingName);

const addressValue = document.getElementById("app-address-value");
addressValue.addEventListener("focusout", () => stopEditAddress(false));
addressValue.addEventListener("keyup", e => {
    if (e.code === "Escape") {
        stopEditAddress(false);
        return;
    }

    if (e.key === "Enter") {
        stopEditAddress(true);
        return;
    }
});

const appIcon = document.getElementById("app-icon");
appIcon.onmousemove = () => showAppInstanceMenu();

document.getElementById("app-add-to-favorites").onclick = async () => {
    await dotnet.toggleCurrentFavorite();
    updateFavoritesButtons();
};

document.getElementById("app-remove-from-favorites").onclick = async () => {
    await dotnet.toggleCurrentFavorite();
    updateFavoritesButtons();
};

const appSort = document.getElementById("app-sort");
appSort.onmousemove = () => showAppSort();

const appActions = document.getElementById("app-actions");
appActions.onmousemove = () => showAppActions();

const appView = document.getElementById("app-view")
appView.onmousemove = () => showAppViewMenu();

document.addEventListener("mousemove", (e) => {
    clearTimeout(tooltipId);
    const elem = document.elementFromPoint(e.clientX, e.clientY);
    if (elem) {
        let tt = getTooltip(elem);
        if (tt) {
            tt = Tools.Resource(tt); // check for localization
            tooltipId = setTimeout(() => {
                tooltip.style.display = "block";
                tooltip.innerHTML = tt;
                Tools.placeElementAtCursor(tooltip, e.clientX, e.clientY, tooltipOffset, tooltipOffset);
            }, 500);
            return;
        }
    }

    tooltip.style.display = "none";
    tooltip.innerHTML = "";
});

function openMenu(id, selectPath) {
    switch (id) {
        case Enums.MenuId.AppView:
            showAppViewMenu(selectPath);
            break;

        case Enums.MenuId.AppInfo:
            showAppInfoMenu(selectPath);
            break;

        case Enums.MenuId.AppInstance:
            showAppInstanceMenu(selectPath);
            break;

        case Enums.MenuId.AppThumbnailSize:
            showAppThumnbailSizeMenu(selectPath);
            break;

        case Enums.MenuId.AppSort:
            showAppSort(selectPath);
            break;

        case Enums.MenuId.AppActions:
            showAppActions(selectPath);
            break;

        case Enums.MenuId.AppFavorites:
            showFavorites(selectPath);
            break;

        case Enums.MenuId.AppHistory:
            showHistory(selectPath);
            break;
    }
}

async function showAppActions(selectPath) {
    const rc = appActions.getBoundingClientRect();
    const actions = await dotnet.getActions();
    const isNewFolderSupported = await actions.isNewFolderSupported;
    const openFromExplorerList = await actions.openFromExplorerList;
    let openFromExplorerItems = [];
    if (openFromExplorerList.length > 0) {
        for (let i = 0; i < openFromExplorerList.length; i++) {
            const item = openFromExplorerList[i];
            openFromExplorerItems.push({
                parsingName: await item.key,
                html: await item.value
            });
        }
    }

    const terminals = await actions.getTerminals();
    let terminalItems = [];
    if (terminals.length > 0) {
        for (let i = 0; i < terminals.length; i++) {
            const item = terminals[i];
            terminalItems.push({
                key: await item.key,
                html: await item.displayName,
                commandLine: await item.commandLine,
                icon: await item.icon,
                tooltip: await item.commandLine
            });
        }
    }

    const openFromVisualStudioList = await actions.getOpenFromVisualStudioList();
    let openFromVisualStudioItems = [];
    if (openFromVisualStudioList.length > 0) {
        for (let i = 0; i < openFromVisualStudioList.length; i++) {
            const item = openFromVisualStudioList[i];
            openFromVisualStudioItems.push({
                parsingName: await item.key,
                html: await item.value
            });
        }
    }

    const detectsVisualStudioInstances = await actions.detectsVisualStudioInstances;
    const menuItems = await actions.GetEntryActions();
    let items = await Menu.Menu.getItems(menuItems, actions);
    items.push(
        {
            html: Tools.Resource("NewFolder"),
            icon: "fa-solid fa-folder-plus",
            isHidden: !isNewFolderSupported,
            onclick: () => dotnet.createNewFolder()
        },
        {
            html: Tools.Resource("FindStrings"),
            icon: "fa-solid fa-filter",
            tooltip: Tools.Resource("FindStringsDescription"),
            onclick: () => openSearch(Enums.SearchType.FindStrings)
        },
        {
            html: Tools.Resource("WindowsSearch"),
            icon: "fa-solid fa-magnifying-glass",
            tooltip: Tools.Resource("WindowsSearchDescription"),
            onclick: () => openSearch(Enums.SearchType.WindowsSearch)
        },
        { isSeparator: true },
        {
            html: Tools.Resource("ExportAsCsv"),
            icon: "fa-solid fa-file-csv",
            eventType: Enums.WebEventType.ExportAsCsv
        },
        { isSeparator: true, isHidden: openFromExplorerList.length == 0 },
        {
            html: Tools.Resource("OpenFromExplorer"),
            icon: "fa-brands fa-windows",
            isHidden: openFromExplorerItems.length == 0,
            items: openFromExplorerItems
        },
        { isSeparator: true, isHidden: !detectsVisualStudioInstances },
        {
            html: Tools.Resource("OpenFromVisualStudio"),
            icon: "fa-solid fa-infinity",
            isHidden: !detectsVisualStudioInstances,
            items: openFromVisualStudioItems
        },
        { isSeparator: true, isHidden: terminalItems.length == 0 },
        {
            html: Tools.Resource("NewTerminal"),
            icon: "fa-solid fa-terminal",
            isHidden: terminalItems.length == 0,
            items: terminalItems
        }
    );

    if (!window.isAdministrator) {
        items.push({ isSeparator: true });
        items.push({
            html: Tools.Resource("RestartAsAdministrator"),
            icon: "fa-solid fa-shield-halved",
            eventType: Enums.WebEventType.RestartAsAdministrator
        });
    }

    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: items
        },
        options: {
            id: Enums.MenuId.AppActions,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            subMenuOffsetX: rc.height,
            animations: {
                duration: ".4s",
                show: "fadeInUp",
                hide: "fadeOutDown"
            },
            onclick: (e, item) => {
                if (!item)
                    return;

                if (item.commandLine) {
                    dotnet.runTerminal(item.key);
                    return;
                }

                if (item.parsingName) {
                    navigate(item.parsingName);
                    return;
                }

                if (item.eventType) {
                    dotnet.sendEvent(item.eventType);
                    e.preventDefault();
                    e.stopPropagation();
                    return;
                }
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function showAppSort(selectPath) {
    const rc = appSort.getBoundingClientRect();
    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: [
                {
                    html: Tools.Resource("Name"),
                    sortBy: Enums.SortBy.Name,
                    isChecked: window.sortBy == Enums.SortBy.Name
                },
                {
                    html: Tools.Resource("DateModified"),
                    sortBy: Enums.SortBy.DateModified,
                    isChecked: window.sortBy == Enums.SortBy.DateModified
                },
                {
                    html: Tools.Resource("Size"),
                    sortBy: Enums.SortBy.Size,
                    isChecked: window.sortBy == Enums.SortBy.Size
                },
                {
                    html: Tools.Resource("Extension"),
                    sortBy: Enums.SortBy.Extension,
                    isChecked: window.sortBy == Enums.SortBy.Extension
                },
                { isSeparator: true },
                {
                    html: Tools.Resource("Ascending"),
                    sortDirection: Enums.SortDirection.Ascending,
                    isChecked: window.sortDirection == Enums.SortDirection.Ascending
                },
                {
                    html: Tools.Resource("Descending"),
                    sortDirection: Enums.SortDirection.Descending,
                    isChecked: window.sortDirection == Enums.SortDirection.Descending
                }
            ]
        },
        options: {
            id: Enums.MenuId.AppSort,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            subMenuOffsetX: rc.height,
            animations: {
                duration: ".4s",
                show: "fadeInUp",
                hide: "fadeOutDown"
            },
            onclick: (e, item) => {
                if (item.sortBy) {
                    window.sortBy = item.sortBy;
                    updateEntries();
                    dotnet.setInstanceSetting("sortBy", window.sortBy);
                }
                else if (item.sortDirection) {
                    window.sortDirection = item.sortDirection;
                    updateEntries();
                    dotnet.setInstanceSetting("sortDirection", window.sortDirection);
                }
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function showAppThumnbailSizeMenu(selectPath) {
    const rc = appThumbnailsSize.getBoundingClientRect();
    const settings = JSON.parse(syncDotnet.getSettings());
    const currentSize = window.thumbnailsSize;
    const allowedThumbnailsSizes = settings.allowedThumbnailsSizes;
    if (allowedThumbnailsSizes.find(size => size === currentSize) == null) {
        allowedThumbnailsSizes.push(currentSize);
        allowedThumbnailsSizes.sort((a, b) => a - b);
    }

    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: allowedThumbnailsSizes.map(size => {
                return {
                    html: size == currentSize ? "<div class='thumbnails-size'>" + size + " px</div>" : size + " px",
                };
            })
        },
        options: {
            id: Enums.MenuId.AppThumbnailSize,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            subMenuOffsetX: rc.height,
            animations: {
                duration: ".4s",
                show: "fadeInUp",
                hide: "fadeOutDown"
            },
            onclick: (e, item) => {
                if (item && item.html) {
                    saveThumbnailsSize(parseInt(item.html));
                    updateEntries();
                }
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function showAppInstanceMenu(selectPath) {
    const rc = appIcon.getBoundingClientRect();
    const instances = syncDotnet.getInstances();
    const screens = syncDotnet.getScreens();
    const configuration = {
        menu: {
            items: [
                {
                    html: Tools.Resource("OpenNewInstance"),
                    onclick: () => dotnet.sendEvent(Enums.WebEventType.OpenNewInstance),
                    icon: "fa-solid fa-circle-plus",
                },
                {
                    html: Tools.Resource("OpenNewInstanceAdministrator"),
                    onclick: () => dotnet.sendEvent(Enums.WebEventType.OpenNewInstanceAdministrator),
                    icon: "fa-solid fa-shield-halved",
                    isHidden: window.isAdministrator
                },
                { isSeparator: true, isHidden: instances.length <= 1 },
                {
                    isHidden: instances.length <= 1,
                    html: Tools.Resource("ArrangeInstances"),
                    items: [
                        {
                            html: Tools.Resource("Vertically"),
                            icon: "fa-solid fa-arrows-up-down",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ArrangeInstances, { arrangement: "vertically" })
                        },
                        {
                            html: Tools.Resource("Horizontally"),
                            icon: "fa-solid fa-arrows-left-right",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ArrangeInstances, { arrangement: "horizontally" })
                        },
                        {
                            html: Tools.Resource("Tiled"),
                            icon: "fa-solid fa-border-all",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ArrangeInstances, { arrangement: "tiled" })
                        },
                        { isSeparator: true, isHidden: screens.length <= 1 },
                        {
                            html: Tools.Resource("VerticallyOnAllScreens"),
                            isHidden: screens.length <= 1,
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ArrangeInstances, { arrangement: "vertically", allScreens: true })
                        },
                        {
                            html: Tools.Resource("HorizontallyOnAllScreens"),
                            isHidden: screens.length <= 1,
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ArrangeInstances, { arrangement: "horizontally", allScreens: true })
                        },
                        {
                            html: Tools.Resource("TiledOnAllScreens"),
                            isHidden: screens.length <= 1,
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ArrangeInstances, { arrangement: "tiled", allScreens: true })
                        }
                    ]
                },
                { isSeparator: true, isHidden: instances.length <= 1 },
                {
                    isHidden: instances.length <= 1,
                    html: Tools.Resource("QuitInstance"),
                    items: instances.filter(inst => !inst.IsThis).map(inst => {
                        return {
                            html: inst.otherDisplayName,
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.QuitInstance, { processId: inst.processId })
                        };
                    })
                },
                {
                    isHidden: instances.length <= 1,
                    html: Tools.Resource("QuitAllInstances"),
                    onclick: () => dotnet.sendEvent(Enums.WebEventType.QuitAllInstances)
                }
            ]
        },
        options: {
            id: Enums.MenuId.AppInstance,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            animations: {
                duration: ".4s",
                show: "fadeInUp",
                hide: "fadeOutDown"
            }
        }
    };

    let splicePos = 2;
    if (screens.length > 1) {
        for (let i = 0; i < screens.length; i++) {
            const screen = screens[i];
            if (screen.isThis)
                continue;

            configuration.menu.items.splice(splicePos++, 0, {
                html: Tools.Resource("OpenNewInstanceScreen").replace(/\{0\}/, screen.displayName),
                icon: "fa-solid fa-display",
                onclick: () => dotnet.sendEvent(Enums.WebEventType.OpenNewInstanceOnScreen, { devicePath: screen.DevicePath }),
            });
        }
    }

    if (instances.length > 1) {
        let instancesItems = [];
        for (let i = 0; i < instances.length; i++) {
            const instance = instances[i];
            if (instance.isThis)
                continue;

            instancesItems.push({
                html: instance.otherDisplayName,
                onclick: () => dotnet.sendEvent(Enums.WebEventType.SwitchToInstance, { processId: instance.processId })
            });
        }

        configuration.menu.items.splice(splicePos++, 0, { isSeparator: true });
        configuration.menu.items.splice(splicePos, 0, {
            html: Tools.Resource("SwitchToInstance"),
            items: instancesItems
        });
    }

    configuration.menu.items.push({ isSeparator: true });
    configuration.menu.items.push({
        html: Tools.Resource("SaveInstance"),
        icon: "fa-solid fa-floppy-disk",
        onclick: () => dotnet.saveInstance()
    });

    const menu = new Menu.Menu();
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function showAppInfoMenu(selectPath) {
    const rc = appInfo.getBoundingClientRect();
    let instanceSettingsName;
    if (window.instanceName && window.instanceName.length > 0) {
        instanceSettingsName = Tools.Resource("NamedInstanceSettings", { name: window.instanceName });
    }
    else {
        instanceSettingsName = Tools.Resource("InstanceSettings");
    }

    let themes = [];
    themes.push({
        html: Tools.Resource("LoadTheme"),
        icon: "fa-solid fa-file-import",
        onclick: () => dotnet.loadTheme()
    });
    themes.push({
        html: Tools.Resource("SaveTheme"),
        icon: "fa-solid fa-palette",
        onclick: () => dotnet.saveTheme()
    });
    themes.push({
        html: Tools.Resource("RefreshCurrentTheme"),
        icon: "fa-solid fa-rotate",
        onclick: () => dotnet.refreshCurrentTheme()
    });
    themes.push({ isSeparator: true });
    themes.push(...syncDotnet.getAvailableThemes().map(th => {
        return {
            html: th.isCurrent ? "<div class='current-theme'>" + th.displayName + "</div>" : th.displayName,
            tooltip: th.filePath,
            onclick: () => dotnet.loadTheme(th.filePath)
        }
    }));

    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: [
                {
                    html: instanceSettingsName + "...",
                    icon: "fa-solid fa-gear",
                    onclick: () => showPropertyGrid(Enums.PropertyGridType.InstanceSettings)
                },
                {
                    html: Tools.Resource("GlobalSettings") + "...",
                    icon: "fa-solid fa-gears",
                    onclick: () => showPropertyGrid(Enums.PropertyGridType.Settings)
                },
                {
                    html: Tools.Resource("KeyboardShortcuts") + "...",
                    icon: "fa-solid fa-keyboard",
                    onclick: () => showPropertyGrid(Enums.PropertyGridType.KeyboardShortcuts)
                },
                { isSeparator: true },
                {
                    html: Tools.Resource("Themes"),
                    icon: "fa-solid fa-palette",
                    items: themes
                },
                { isSeparator: true },
                {
                    html: Tools.Resource("Tools"),
                    icon: "fa-solid fa-toolbox",
                    items: [
                        {
                            html: Tools.Resource("OpenConfig"),
                            icon: "fa-solid fa-user-gear",
                            onclick: () => dotnet.openConfigurationFolder()
                        },
                        {
                            html: Tools.Resource("OpenLogsFolder"),
                            icon: "fa-solid fa-stethoscope",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.OpenLogsFolder)
                        },
                        { isSeparator: true },
                        {
                            html: Tools.Resource("ExportExtensionsAsCsv"),
                            icon: "fa-solid fa-file-csv",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.ExportExtensionsAsCsv)
                        },
                        { isSeparator: true },
                        {
                            html: Tools.Resource("RegisterShellIntegrations"),
                            icon: "fa-brands fa-windows",
                            onclick: () => dotnet.registerShellIntegrations()
                        },
                        {
                            html: Tools.Resource("UnregisterShellIntegrations"),
                            icon: "fa-brands fa-windows",
                            onclick: () => dotnet.unregisterShellIntegrations()
                        },
                        { isSeparator: true },
                        {
                            html: Tools.Resource("RemoveDeletedItemsFromHistory"),
                            icon: "fa-solid fa-file-circle-xmark",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.RemoveDeletedItemsFromHistory)
                        },
                        {
                            html: Tools.Resource("RemoveDeletedItemsFromFavorites"),
                            icon: "fa-solid fa-heart-circle-xmark",
                            onclick: () => dotnet.sendEvent(Enums.WebEventType.RemoveDeletedItemsFromFavorites)
                        },
                        {
                            html: Tools.Resource("DeleteHistory"),
                            icon: "fa-solid fa-timeline",
                            onclick: () => dotnet.deleteHistory()
                        },
                        {
                            html: Tools.Resource("ClearCaches"),
                            icon: "fa-solid fa-square-binary",
                            onclick: () => dotnet.clearCaches()
                        },
                        {
                            isHidden: !window.showDevTools,
                            html: "GC Collect",
                            icon: "fa-solid fa-skull",
                            onclick: () => dotnet.gcCollect()
                        },
                    ]
                },
                { isSeparator: true },
                {
                    html: Tools.Resource("CheckForUpdates") + "...",
                    icon: "fa-solid fa-download",
                    onclick: () => dotnet.checkForUpdates()
                },
                {
                    html: Tools.Resource("SysInfo") + "...",
                    icon: "fa-solid fa-computer",
                    onclick: () => showPropertyGrid(Enums.PropertyGridType.Info)
                },
                {
                    html: Tools.Resource("About") + "...",
                    icon: "fa-solid fa-circle-info",
                    onclick: async () => {
                        const html = await dotnet.getAboutHtml();
                        Swal.fire({
                            title: Tools.Resource("About"),
                            html: html,
                            customClass: "fld-about"
                        });
                    }
                }
            ]
        },
        options: {
            id: Enums.MenuId.AppInfo,
            className: "fld-menu",
            left: rc.left,
            top: rc.bottom,
            subMenuOffsetY: rc.height,
            animations: {
                duration: ".4s",
                show: "fadeInDown",
                hide: "fadeOutUp"
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function showAppViewMenu(selectPath) {
    const rc = appView.getBoundingClientRect();
    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: [
                {
                    html: Tools.Resource("Details"),
                    viewBy: Enums.ViewBy.Details,
                    isChecked: window.viewBy == Enums.ViewBy.Details
                },
                {
                    html: Tools.Resource("Images"),
                    viewBy: Enums.ViewBy.Images,
                    isChecked: window.viewBy == Enums.ViewBy.Images
                },
                { isSeparator: true },
                {
                    html: Tools.Resource("Options"),
                    items: [
                        {
                            html: Tools.Resource("SelectionOptions"),
                            className: "fld-menu-items-group"
                        },
                        {
                            html: Tools.Resource("ExcludeFolders"),
                            entryEnumerateOptions: Enums.EntryEnumerateOptions.ExcludeFolders,
                            isChecked: window.entryEnumerateOptions & Enums.EntryEnumerateOptions.ExcludeFolders
                        },
                        {
                            html: Tools.Resource("ExcludeFiles"),
                            entryEnumerateOptions: Enums.EntryEnumerateOptions.ExcludeFiles,
                            isChecked: window.entryEnumerateOptions & Enums.EntryEnumerateOptions.ExcludeFiles
                        },
                        {
                            html: Tools.Resource("IncludeHidden"),
                            entryEnumerateOptions: Enums.EntryEnumerateOptions.IncludeHidden,
                            isChecked: window.entryEnumerateOptions & Enums.EntryEnumerateOptions.IncludeHidden
                        },
                        {
                            html: Tools.Resource("IncludeSystem"),
                            entryEnumerateOptions: Enums.EntryEnumerateOptions.IncludeSystem,
                            isChecked: window.entryEnumerateOptions & Enums.EntryEnumerateOptions.IncludeSystem
                        },
                        {
                            html: Tools.Resource("ShowCompressedAsFile"),
                            entryEnumerateOptions: Enums.EntryEnumerateOptions.ShowCompressedAsFile,
                            isChecked: window.entryEnumerateOptions & Enums.EntryEnumerateOptions.ShowCompressedAsFile
                        },
                        {
                            html: Tools.Resource("DetailsOptions"),
                            className: "fld-menu-items-group"
                        },
                        {
                            html: Tools.Resource("ShowIcons"),
                            viewByDetailsOptions: Enums.ViewByDetailsOptions.ShowIcons,
                            isChecked: window.viewByDetailsOptions & Enums.ViewByDetailsOptions.ShowIcons
                        },
                        {
                            html: Tools.Resource("ShowThumbnails"),
                            viewByDetailsOptions: Enums.ViewByDetailsOptions.ShowThumbnails,
                            isChecked: window.viewByDetailsOptions & Enums.ViewByDetailsOptions.ShowThumbnails
                        },
                        {
                            html: Tools.Resource("ImagesOptions"),
                            className: "fld-menu-items-group"
                        },
                        {
                            html: Tools.Resource("ShowImagesOnly"),
                            entryEnumerateOptions: Enums.EntryEnumerateOptions.ImagesOnly,
                            isChecked: window.entryEnumerateOptions & Enums.EntryEnumerateOptions.ImagesOnly
                        },
                        {
                            html: Tools.Resource("DisplayTitle"),
                            viewByImageOptions: Enums.ViewByImageOptions.DisplayTitle,
                            isChecked: window.viewByImageOptions & Enums.ViewByImageOptions.DisplayTitle
                        },
                        {
                            html: Tools.Resource("RenderPdfThumbnails"),
                            viewByImageOptions: Enums.ViewByImageOptions.RenderPdfThumbnails,
                            isChecked: window.viewByImageOptions & Enums.ViewByImageOptions.RenderPdfThumbnails
                        },
                        {
                            html: Tools.Resource("SquareThumbnails"),
                            viewByImageOptions: Enums.ViewByImageOptions.SquareThumbnails,
                            isChecked: window.viewByImageOptions & Enums.ViewByImageOptions.SquareThumbnails
                        }
                    ]
                }
            ]
        },
        options: {
            id: Enums.MenuId.AppView,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            subMenuOffsetX: rc.height,
            animations: {
                duration: ".4s",
                show: "fadeInUp",
                hide: "fadeOutDown"
            },
            onclick: (e, item) => {
                if (!item)
                    return;

                if (item.viewBy) {
                    window.viewBy = item.viewBy;
                    updateThumbnailsSizeControls();
                    updateEntries();
                    dotnet.setInstanceSetting("viewBy", window.viewBy);
                }
                else if (item.viewByImageOptions) {
                    if (item.viewByImageOptions == Enums.ViewByImageOptions.DisplayTitle) {
                        if (window.viewByImageOptions & Enums.ViewByImageOptions.DisplayTitle) {
                            window.viewByImageOptions &= ~Enums.ViewByImageOptions.DisplayTitle;
                        }
                        else {

                            window.viewByImageOptions |= Enums.ViewByImageOptions.DisplayTitle;
                        }
                    }
                    else if (item.viewByImageOptions == Enums.ViewByImageOptions.RenderPdfThumbnails) {
                        if (window.viewByImageOptions & Enums.ViewByImageOptions.RenderPdfThumbnails) {
                            window.viewByImageOptions &= ~Enums.ViewByImageOptions.RenderPdfThumbnails;
                        }
                        else {
                            window.viewByImageOptions |= Enums.ViewByImageOptions.RenderPdfThumbnails;
                        }
                    }
                    else if (item.viewByImageOptions == Enums.ViewByImageOptions.SquareThumbnails) {
                        if (window.viewByImageOptions & Enums.ViewByImageOptions.SquareThumbnails) {
                            window.viewByImageOptions &= ~Enums.ViewByImageOptions.SquareThumbnails;
                        }
                        else {
                            window.viewByImageOptions |= Enums.ViewByImageOptions.SquareThumbnails;
                        }
                    }

                    updateEntries();
                    dotnet.setInstanceSetting("viewByImageOptions", window.viewByImageOptions);
                }
                else if (item.entryEnumerateOptions) {
                    // can't have both exclude folders and exclude files
                    if (item.entryEnumerateOptions == Enums.EntryEnumerateOptions.ExcludeFolders &&
                        window.entryEnumerateOptions & Enums.EntryEnumerateOptions.ExcludeFiles) {
                        window.entryEnumerateOptions &= ~Enums.EntryEnumerateOptions.ExcludeFiles;
                    }
                    else if (item.entryEnumerateOptions == Enums.EntryEnumerateOptions.ExcludeFiles &&
                        window.entryEnumerateOptions & Enums.EntryEnumerateOptions.ExcludeFolders) {
                        window.entryEnumerateOptions &= ~Enums.EntryEnumerateOptions.ExcludeFolders;
                    }

                    window.entryEnumerateOptions = Tools.toggleFlags(window.entryEnumerateOptions, item.entryEnumerateOptions);
                    updateEntries();
                    dotnet.setInstanceSetting("entryEnumerateOptions", window.entryEnumerateOptions);
                }
                else if (item.viewByDetailsOptions) {
                    // can't have both icons & thumbnails
                    if (item.viewByDetailsOptions == Enums.ViewByDetailsOptions.ShowIcons &&
                        window.viewByDetailsOptions & Enums.ViewByDetailsOptions.ShowThumbnails) {
                        window.viewByDetailsOptions &= ~Enums.ViewByDetailsOptions.ShowThumbnails;
                    }
                    else if (item.viewByDetailsOptions == Enums.ViewByDetailsOptions.ShowThumbnails &&
                        window.viewByDetailsOptions & Enums.ViewByDetailsOptions.ShowIcons) {
                        window.viewByDetailsOptions &= ~Enums.ViewByDetailsOptions.ShowIcons;
                    }

                    window.viewByDetailsOptions = Tools.toggleFlags(window.viewByDetailsOptions, item.viewByDetailsOptions);
                    updateEntries();
                    dotnet.setInstanceSetting("viewByDetailsOptions", window.viewByDetailsOptions);
                }
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function openSearch(type) {
    new Search.Search({ type: type })
}

function applyTheme(theme) {
    for (const key in theme) {
        if (key == "css-file-name") {
            const cssPath = theme[key];
            if (cssPath) {
                let link = document.querySelector("link[href='" + cssPath + "']");
                if (!link) {
                    link = document.createElement("link");
                    link.rel = "stylesheet";
                    link.href = cssPath;
                    link.shellBatTheme = true;
                    document.head.appendChild(link);
                }
            }
            else {
                const links = document.querySelectorAll("link[rel='stylesheet']");
                links.forEach(link => {
                    if (link.shellBatTheme) {
                        document.head.removeChild(link);
                    }
                });
            }

            continue;
        }
        document.documentElement.style.setProperty(`--${key}`, theme[key]);
    }
}

function updateViewFilter(e) {
    if (e.key.length != 1 || e.shiftKey || e.ctrlKey || e.metaKey || e.altKey)
        return false;

    filterDiv.style.visibility = "visible";
    viewFilterInput.value = e.key;
    viewFilterInput.focus();
    filterEntries();
    highlightViewFilterText();
    return true;
}

function highlightViewFilterText() {
    const filter = viewFilterInput.value.trim().toLowerCase();
    if (filter.length == 0)
        return;

    const entriesElement = document.getElementById("app-entries");
    if (entriesElement && entriesElement.rows) {
        if (entriesElement.rows.length == 0)
            return;

        const prc = entriesParent.getBoundingClientRect();
        const trc = entriesElement.getBoundingClientRect(); // get first visible row
        const firstCell = document.elementsFromPoint(trc.left + 1, prc.top + 1).find(e => e.tagName === "TD");
        let row;
        if (!firstCell) {
            row = entriesElement.rows[0];
        }
        else {
            row = firstCell.parentElement;
        }

        do {
            const rrc = row.getBoundingClientRect();
            if (rrc.top > prc.bottom)
                break;

            const cell = row.cells[3]; // Name column
            if (cell) {
                const text = cell.innerText.toLowerCase();
                const index = text.indexOf(filter);
                if (index >= 0) {
                    const highlightedText = cell.innerText.substring(0, index) +
                        "<mark>" +
                        cell.innerText.substring(index, index + filter.length) +
                        "</mark>" +
                        cell.innerText.substring(index + filter.length);
                    cell.innerHTML = highlightedText;
                }
                else {
                    cell.innerHTML = cell.innerText;
                }
            }

            row = row.nextElementSibling;
        }
        while (row);
    } else { // images
        const displayTitle = window.viewByImageOptions & Enums.ViewByImageOptions.DisplayTitle;
        if (!displayTitle)
            return;

        const imagesElement = document.getElementById("app-images");
        const prc = entriesParent.getBoundingClientRect();
        const trc = imagesElement.getBoundingClientRect();  // get first visible div
        let firstDiv = document.elementsFromPoint(trc.left + 1, prc.top + 1).find(e => e.parentElement == imagesElement && e.tagName === "DIV");
        if (!firstDiv) {
            firstDiv = imagesElement.firstElementChild;
        }

        do {
            const rrc = firstDiv.getBoundingClientRect();
            if (rrc.top > prc.bottom)
                break;

            const titleElement = firstDiv.querySelector(".app-image-title");
            const text = titleElement.innerText.toLowerCase();
            const index = text.indexOf(filter);
            if (index >= 0) {
                const highlightedText = titleElement.innerText.substring(0, index) +
                    "<mark>" +
                    titleElement.innerText.substring(index, index + filter.length) +
                    "</mark>" +
                    titleElement.innerText.substring(index + filter.length);
                titleElement.innerHTML = highlightedText;
            }
            else {
                titleElement.innerHTML = titleElement.innerText;
            }

            firstDiv = firstDiv.nextElementSibling;
        }
        while (firstDiv);
    }
}

function clearViewFilter() {
    filterDiv.style.visibility = "collapse";
    viewFilterInput.value = "";
    filterEntries();
}

function showHistory(selectPath) {
    const rc = appHistory.getBoundingClientRect();
    const list = syncDotnet.getHistory();
    if (list.length == 0)
        return;

    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: list.map(h => {
                return {
                    html: h.displayName ?? h.parsingName,
                    parsingName: h.parsingName,
                    iconPath: h.iconPath,
                    className: h.parsingName == window.parsingName ? "current" : null,
                    onDelete: () => {
                        syncDotnet.removeHistoryEntries(h.parsingName);
                        Menu.Menu.dismissRootMenu();
                        showHistory();
                    }
                };
            })
        },
        options: {
            id: Enums.MenuId.AppHistory,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            animations: {
                duration: ".4s",
                show: "fadeInDown",
                hide: "fadeOutUp"
            },
            onclick: (e, item) => {
                if (item) {
                    navigate(item.parsingName);
                }
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

async function updateFavoritesButtons() {
    const isFavorite = await dotnet.isCurrentFavorite();
    document.getElementById("app-add-to-favorites").style.display = isFavorite ? "none" : "block";
    document.getElementById("app-remove-from-favorites").style.display = isFavorite ? "block" : "none";
}

function showFavorites(selectPath) {
    const rc = appFavorites.getBoundingClientRect();
    const list = syncDotnet.getFavorites();
    if (list.length == 0)
        return;

    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: list.map(h => {
                return {
                    html: h.displayName ?? h.parsingName,
                    parsingName: h.parsingName,
                    iconPath: h.iconPath,
                    className: h.parsingName == window.parsingName ? "current" : null
                };
            })
        },
        options: {
            id: Enums.MenuId.AppFavorites,
            className: "fld-menu",
            left: rc.left,
            top: rc.top,
            animations: {
                duration: ".4s",
                show: "fadeInDown",
                hide: "fadeOutUp"
            },
            onclick: (e, item) => {
                if (item) {
                    navigate(item.parsingName);
                }
            }
        }
    };
    configuration.selectPath = selectPath;
    menu.draw(configuration);
}

function saveThumbnailsSize(size) {
    window.thumbnailsSize = size;
    dotnet.setInstanceSetting("ThumbnailsSize", window.thumbnailsSize);
}

function setThumbnailsSize(size) {
    Menu.Menu.dismissRootMenu();
    saveThumbnailsSize(size);
    updateThumbnailsSizeControls();
    updateEntries();
}

function updateThumbnailsSizeControls() {
    if (window.viewBy != Enums.ViewBy.Details) {
        appThumbnailsSize.style.display = "block";
        appThumbnailsSize.classList.remove("animate__bounceOutDown");
        appThumbnailsSize.classList.add("animate__bounceInUp");
    }
    else {
        appThumbnailsSize.classList.remove("animate__bounceInUp");
        appThumbnailsSize.classList.add("animate__bounceOutDown");
        appThumbnailsSize.addEventListener('animationend', () => {
            appThumbnailsSize.style.display = "none";
        }, { once: true });
    }
}

function getSelection() {
    let parsingNames = [];

    if (selection && Object.keys(selection).length > 0) {
        for (let key in selection) {
            const element = document.getElementById("e#" + key);
            if (element && element.parsingName) {
                parsingNames.push(element.parsingName);
            }
        }
    }

    return parsingNames;
}

function getParsingNameAndSelectionAtPoint(x, y) {
    const element = document.elementFromPoint(x, y);
    if (!element)
        return null;

    let entry = Tools.findParent(element, "TR"); // details
    if (!entry || !entry.parsingName) {
        entry = Tools.findParent(element, "DIV"); // images
        if (!entry || !entry.parsingName)
            return null;
    }

    let parsingNames = [entry.parsingName];

    if (selection && Object.keys(selection).length > 0) {
        for (let key in selection) {
            const element = document.getElementById("e#" + key);
            if (element && element.parsingName) {
                parsingNames.push(element.parsingName);
            }
        }
    }

    return parsingNames;
}

function getEntriesRect() {
    const rc = entriesParent.getBoundingClientRect();
    const json = {
        left: rc.left,
        top: rc.top,
        right: rc.right,
        bottom: rc.height,
    };
    return json;
}

function showTerminal(id, key, options) {
    new Terminus.Terminus(id, key, options);
}

function closeTerminal(id) {
    Terminus.Terminus.close(id);
}

function goBack() {
    dotnet.sendEvent(Enums.WebEventType.MoveHistoryBack);
}

function goForward() {
    dotnet.sendEvent(Enums.WebEventType.MoveHistoryForward);
}

async function showWindow(name, parameters, options) {
    const w = await dotnet.getWindow(name, parameters);
    if (!w)
        return;

    options = options || {};
    const newWindow = options.newWindow || false;
    const forceOpen = options.forceOpen || options.newWindow || false;
    const id = await w.id;
    let win = newWindow ? undefined : Window.Window.get(id);
    if (win) {
        win.update(w);
        win.setPosition(options);
        if (options.viewerId) {
            const module = await win.selectViewerModule(options.viewerId);
            if (module && options.viewerOptions) {
                module.setOptions(options.viewerOptions);
            }
        }
    }
    else if (forceOpen) {
        if (window.isShowing)
            return;

        window.isShowing = true;
        win = new Window.Window(w);
        win.addEventListener("ready", async () => {
            win.setPosition(options);
            if (options.viewerId) {
                const module = await win.selectViewerModule(options.viewerId);
                if (module && options.viewerOptions) {
                    module.setOptions(options.viewerOptions);
                }
            }
            delete window.isShowing;
        }, { once: true });
    } else {
        // send an event to .NET to notify that the window is not open
        dotnet.sendEvent(Enums.WebEventType.WindowNotOpen, { name: name, parsingName: parameters, options: options });
    }
}

async function showPropertyGrid(type) {
    const wpg = await dotnet.getPropertyGrid(type);
    const pg = new PropertyGrid.PropertyGrid();
    const div = document.createElement("div");
    div.className = "fld-pg-container";
    await pg.draw(div, wpg);

    const options = await wpg.options;
    const title = await options.title || Tools.Resource("Properties");
    const swalClassName = await options.swalClassName;
    window.swalModal = true;
    Swal.fire({
        title: title,
        html: div,
        customClass: swalClassName
    }).then((e) => {
        delete window.swalModal;
        if (e.isConfirmed) {
            wpg.save();
        }
    });
}

async function updateInstances() {
    const instances = await dotnet.getInstances();
    const div = document.getElementById("app-instances");
    const thisInstance = instances.find(i => i.isThis);
    const name = await thisInstance.displayName;
    if (instances.length <= 1) {
        div.innerHTML = name;
    }
    else {
        div.innerHTML = name + " (" + instances.length + ")";
    }
}

function editAddress() {
    isEditing = true;
    const crumbs = document.getElementById("app-breadcrumbs");
    crumbs.style.display = "none";

    const div = document.getElementById("app-address");
    div.style.display = "block";
    div.style.flexGrow = "100";
    addressValue.value = window.editName || "";
    sendCaptionSizeChanged();
    setTimeout(() => {
        addressValue.select();
        addressValue.focus();
    }, 100);
}

function stopEditAddress(commit) {
    isEditing = false;
    const crumbs = document.getElementById("app-breadcrumbs");
    crumbs.style.display = "flex";

    const div = document.getElementById("app-address");
    div.style.display = "none";
    div.style.flexGrow = null;
    sendCaptionSizeChanged();
    if (commit) {
        navigate(addressValue.value);
    }
}

function getTooltip(element) {
    if (!element) return null;
    let tt = null;
    if ('getAttribute' in element) {
        tt = element.getAttribute("tooltip");
        if (tt) return tt;
    }
    tt = element.parsingName; // some elements have parsingName property wich holds tooltip
    if (tt) return tt;
    return getTooltip(element.parentNode);
}

function resize() {
    Window.Window.moveSnappedWindows();
    Menu.Menu.dismissRootMenu();
    sendCaptionSizeChanged();
}

function sendCaptionSizeChanged() {
    const rct = document.querySelector("#app-toolbar").getBoundingClientRect();
    const rc = document.querySelector("#app-caption").getBoundingClientRect();
    const rcz = document.querySelector("#app-zoom").getBoundingClientRect()
    const sumRc = {
        x: rc.left,
        y: rc.top,
        width: rc.width + rcz.width,
        height: rct.height
    };
    dotnet.sendEvent(Enums.WebEventType.CaptionSizeChanged, sumRc);
}

function getEntryElement(e) {
    if (e.type == "keydown") {
        return document.getElementById("e#" + selectionIndex);
    }

    let element = Tools.findParent(e.target, "TR"); // details
    if (!element || !element.parsingName) {
        element = Tools.findParent(e.target, "DIV"); // images
        if (!element || !element.parsingName)
            return null;
    }
    return element;
}

function stopRenameEntry(input, parsingName) {
    if (parsingName) { // Enter
        const error = syncDotnet.renameEntry(parsingName, input.value);
        if (error) {
            showAlert({ text: error });
            return;
        }
    }

    if (!isEditing)
        return;

    isEditing = false;
    input.remove();
    input.element.lastChild.style.display = "block";
}

async function renameEntry() {
    if (!selection || Object.keys(selection).length != 1)
        return;

    const element = document.getElementById("e#" + selectionIndex);
    if (!element || !element.parsingName)
        return;

    const name = await dotnet.getEntryEditName(element.parsingName);
    if (!name)
        return;

    isEditing = true;
    element.lastChild.style.display = "none"; // hide current name

    const input = document.createElement("input");
    input.element = element;
    input.originalValue = name;
    input.className = "app-rename";
    input.type = "text";
    input.value = name;
    input.addEventListener("focusout", () => {
        stopRenameEntry(input);
    });

    input.addEventListener("keyup", e => {
        if (e.code === "Escape") {
            stopRenameEntry(input);
            return;
        }

        if (e.key === "Enter") {
            stopRenameEntry(input, element.parsingName);
            return;
        }
    });

    element.appendChild(input);

    setTimeout(() => {
        input.select();
        input.focus();
    }, 100);
}

function deleteSelectedEntries() {
    let parsingNames = [];
    for (let key in selection) {
        const element = document.getElementById("e#" + key);
        if (element && element.parsingName) {
            parsingNames.push(element.parsingName);
        }
    }
    if (parsingNames.length == 0)
        return;

    dotnet.executeAction("Delete", { parsingNames: parsingNames });
    resetSelection();
}

async function entryContextMenu(e) {
    const isEntry = entryClicked(e);

    let parsingNames = [];
    if (isEntry) {
        for (let key in selection) {
            const element = document.getElementById("e#" + key);
            if (element && element.parsingName) {
                parsingNames.push(element.parsingName);
            }
        }
    }

    const cm = await dotnet.getContextMenu({
        parsingNames: parsingNames
    });
    if (!cm)
        return;

    const menuItems = await cm.menuItems;
    if (menuItems.length == 0)
        return;

    let left;
    let top;
    if (!e.clientX || !e.clientY) {
        const element = document.getElementById("e#" + selectionIndex);
        if (element) {
            const rect = element.getBoundingClientRect();
            left = rect.left + rect.width / 2;
            top = rect.bottom - rect.height / 2;
        }
    }
    else {
        left = e.clientX;
        top = e.clientY;
    }

    let items = await Menu.Menu.getItems(menuItems, cm);

    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: items
        },
        options: {
            id: "app-context-menu" + Date.now(),
            className: "fld-menu",
            left: left,
            top: top,
            animations: {
                duration: ".4s",
                show: "fadeIn",
                hide: "fadeOut"
            },
            onclick: (e, item) => {
                if (!item)
                    return;
            }
        }
    };
    menu.draw(configuration);

    if (cm.preventDefault) {
        e.preventDefault();
    }
}

function entryClicked(e) {
    const element = getEntryElement(e);
    if (!element)
        return false;

    const index = parseInt(element.id.substring(2)); // e#
    if (index < 0)
        return false;

    selectEntry(index, selectionIndex, e.shiftKey, e.ctrlKey, e.button === 2);
    selectionIndex = index;
    return true;
}

function entryDoubleClicked(e) {
    const element = getEntryElement(e);
    if (!element)
        return;

    if (element.isFolder || element.isCompressed) {
        navigate(element.parsingName);
        return;
    }

    dotnet.sendEvent(Enums.WebEventType.EntryDoubleClicked, { parsingName: element.parsingName });
}

function showToast(text, duration) {
    Toastify({
        text: text,
        duration: duration || 3000,
        newWindow: true,
        className: "fld-toast",
        close: true,
        gravity: "botttom",
        position: "right",
        stopOnFocus: true,
        offset: {
            x: 14,
            y: 20
        }
    }).showToast();
}

function showAlert(options) {
    if (window.swalModal)
        return;

    Swal.fire({
        position: options?.position || "bottom",
        width: options?.width,
        toast: options?.toast || true,
        icon: options?.icon,
        title: options?.title,
        text: options?.text,
        showConfirmButton: options?.showConfirmButton || false,
        showDenyButton: options?.showDenyButton || false,
        showCancelButton: options?.showCancelButton || false,
        showCloseButton: options?.showCloseButton || false,
        reverseButtons: options?.reverseButtons || false,
        confirmButtonText: options?.confirmButtonText || Tools.Resource("OK"),
        denyButtonText: options?.denyButtonText || Tools.Resource("No"),
        cancelButtonText: options?.cancelButtonText || Tools.Resource("Cancel"),
        timer: options?.timer,
        timerProgressBar: options?.timerProgressBar || false,
        customClass: options?.customClass || "fld-swal"
    }).then(result => {
        if (result.isConfirmed && options.confirmEventType) {
            dotnet.sendEvent(options.confirmEventType);
        }
    });
}

function navigate(parsingName, fromHistory, selectParsingName) {
    const history = syncDotnet.getFirstHistoryEntry();
    if (parsingName !== "") { // desktop is ""
        parsingName = parsingName || (history ? history.parsingName : null);
    }

    const folder = syncDotnet.getFolder(parsingName, fromHistory);
    if (!folder) {
        if (parsingName == null || parsingName == "")
            return;

        showToast(Tools.Resource("CannotNavigate").replace(/\{0\}/, parsingName));
        return;
    }

    resetSelection();

    const oldParsingName = window.parsingName;
    window.parsingName = parsingName;
    window.folder = folder;
    window.editName = folder.editName;
    window.upFullDisplayName = folder.upFullDisplayName;
    window.upParsingName = folder.upParsingName;
    window.backFullDisplayName = folder.backFullDisplayName;
    window.forwardFullDisplayName = folder.forwardFullDisplayName;

    // handle go up tooltip with desktop
    const goUp = document.getElementById("app-go-up");
    goUp.disabled = parsingName == null || parsingName == "";
    goUp.setAttribute("tooltip", goUp.disabled ? "" : "GoUpTo");

    const goBack = document.getElementById("app-go-back");
    goBack.disabled = window.backFullDisplayName == null || window.backFullDisplayName == "";
    goBack.setAttribute("tooltip", goBack.disabled ? "" : "GoBackTo");

    const goForward = document.getElementById("app-go-forward");
    goForward.disabled = window.forwardFullDisplayName == null || window.forwardFullDisplayName == "";
    goForward.setAttribute("tooltip", goForward.disabled ? "" : "GoForwardTo");

    setTimeout(() => {
        updateEntries(folder);
        if (selectParsingName) {
            selectParsingNameEntry(selectParsingName);
        }
        else if (oldParsingName) {
            selectParsingNameEntry(oldParsingName);
        }

        updateFavoritesButtons();
    }, 0);
    setTimeout(() => updateBreadcrumbs(folder), 0);
    sendCaptionSizeChanged();
}

function refreshEntries(parsingName, reason) {
    const folder = window.folder;
    if (!folder)
        return;

    const options = window.entryEnumerateOptions | Enums.EntryEnumerateOptions.DontUseCache;
    setTimeout(() => {
        updateEntries(folder, parsingName, options);
        if (reason) {
            showToast(Tools.Resource("FolderUpdated", { reason: reason }), 3000);
        }
    }, 0);
    setTimeout(() => updateBreadcrumbs(folder), 0);
}

function showBreadcrumbMenu(e, crumb) {
    const rc = e.target.getBoundingClientRect();
    const menu = new Menu.Menu();
    const configuration = {
        menu: {
            items: crumb.getBreadcrumbChildren().map(c => {
                return {
                    html: c.displayName,
                    parsingName: c.parsingName,
                };
            })
        },
        options: {
            id: crumb.parsingName,
            className: "fld-menu",
            left: rc.left,
            top: rc.bottom,
            animations: {
                duration: ".4s",
                show: "fadeInDown",
                hide: "fadeOutUp"
            },
            onclick: (e, item) => {
                if (item) {
                    navigate(item.parsingName);
                }
            }
        }
    };
    if (configuration.menu.items.length == 0)
        return;

    menu.draw(configuration);
}

function updateBreadcrumbs(folder) {
    const crumbs = document.getElementById("app-breadcrumbs");
    crumbs.innerHTML = "";
    const list = folder.entry.getBreadcrumbs();

    // always add desktop first
    let crumbDiv = document.createElement("div");
    crumbDiv.className = "app-crumb-item";
    crumbDiv.innerHTML = "<i class='fa-solid fa-desktop'></i>";
    crumbDiv.setAttribute("tooltip", Tools.Resource("Desktop"));
    crumbDiv.onclick = () => navigate("");
    crumbs.appendChild(crumbDiv);

    for (let i = 1; i < list.length; i++) {
        const c = list[i];

        crumbDiv = document.createElement("div");
        crumbDiv.className = "app-crumb-chevron";
        crumbDiv.setAttribute("tooltip", c.parsingName);
        crumbDiv.onmouseover = (e) => showBreadcrumbMenu(e, c);
        crumbs.appendChild(crumbDiv);

        crumbDiv = document.createElement("div");
        crumbDiv.className = "app-crumb-item";
        crumbDiv.innerHTML = c.displayName;
        crumbDiv.setAttribute("tooltip", c.parsingName);
        crumbDiv.onclick = () => i == list.length - 1 ? editAddress() : navigate(c.parsingName);
        crumbs.appendChild(crumbDiv);
    }

    sendCaptionSizeChanged();
}

function clearSelection() {
    for (const key in selection) {
        document.getElementById("e#" + key)?.classList?.remove("selected");
    }
    resetSelection();
    updateAppItems();
}

function resetSelection() {
    selection = {};
    selectionIndex = -1;
}

function getItemsPerLine() {
    // to be exact, we don't compute bounds, etc. we just scan from index 0 to next line break (or end)
    let index = 0;
    let item = document.getElementById("e#" + index);
    const firstBottom = item.getBoundingClientRect().bottom;
    index++;
    do {
        let item = document.getElementById("e#" + index);
        if (!item)
            return index;

        const itemTop = item.getBoundingClientRect().top;
        if (itemTop >= firstBottom)
            return index;

        index++;
    }
    while (true);
}

function moveSelection(e) {
    if (totalEntries == 0)
        return false;

    if (e.metaKey || e.altKey)
        return false;

    const viewBy = window.viewBy;
    const prevIndex = selectionIndex;
    switch (e.code) {
        case "ArrowUp":
            if (viewBy == Enums.ViewBy.Images) {
                const itemsPerLine = getItemsPerLine();
                if (selectionIndex - itemsPerLine < 0)
                    return true;

                selectionIndex -= itemsPerLine;
            }
            else {
                if (selectionIndex <= 0)
                    return true;

                selectionIndex--;
            }
            break;

        case "ArrowDown":
            if (viewBy == Enums.ViewBy.Images) {
                if (selectionIndex < 0) {
                    selectionIndex = 0;
                }
                else {
                    const itemsPerLine = getItemsPerLine();
                    if (selectionIndex + itemsPerLine >= totalEntries)
                        return true;

                    selectionIndex += itemsPerLine;
                }
            }
            else {
                if (selectionIndex >= totalEntries - 1)
                    return true;

                selectionIndex++;
            }
            break;

        case "ArrowLeft":
            if (viewBy == Enums.ViewBy.Images) {
                if (selectionIndex <= 0)
                    return true;

                selectionIndex--;
            }
            else
                return true;

            break;

        case "ArrowRight":
            if (viewBy == Enums.ViewBy.Images) {
                if (selectionIndex >= totalEntries - 1)
                    return;

                selectionIndex++;
            }
            else
                return true;

            break;

        case "Home":
            if (e.ctrlKey)
                return false;

            selectionIndex = 0;
            break;

        case "End":
            if (e.ctrlKey)
                return false;

            selectionIndex = totalEntries - 1;
            break;

        case "PageUp":
            if (e.ctrlKey)
                return false;

            selectionIndex -= window.paging;
            if (selectionIndex < 0) {
                selectionIndex = 0;
            }
            break;

        case "PageDown":
            if (e.ctrlKey)
                return false;

            selectionIndex += window.paging;
            if (selectionIndex >= totalEntries) {
                selectionIndex = totalEntries - 1;
            }
            break;

        default:
            return false;
    }

    return selectEntry(selectionIndex, prevIndex, e.shiftKey, e.ctrlKey);
}

function selectEntry(index, prevIndex, extend, add, rightClick) {
    const element = document.getElementById("e#" + index);
    if (!element)
        return false;

    // keep current selection if context menu showing and item is already selected
    if (!rightClick || !selection[index]) {
        if (extend) {
            const minIndex = Math.min(prevIndex, index);
            const maxIndex = Math.max(prevIndex, index);

            for (let i = minIndex; i <= maxIndex; i++) {
                selection[i] = true;
                document.getElementById("e#" + i)?.classList?.add("selected");
            }
        }
        else if (add) {
            if (selection[index]) {
                delete selection[index];
                element.classList.remove("selected");
            }
            else {
                selection[index] = true;
                element.classList.add("selected");
            }
        }
        else {
            for (const key in selection) {
                document.getElementById("e#" + key)?.classList?.remove("selected");
            }

            selection = {};
            selection[index] = true;
            element.classList.add("selected");
        }
    }

    Tools.scrollIntoViewIfNeeded(element, entriesParent);

    if (!rightClick) {
        showWindow("Properties", element.parsingName);
    }

    updateAppItems();
    return true;
}

function selectParsingNameEntry(parsingName) {
    if (!parsingName)
        return false;

    for (let i = 0; i < totalEntries; i++) {
        const element = document.getElementById("e#" + i);
        if (element && element.parsingName == parsingName) {
            selectEntry(i, selectionIndex);
            selectionIndex = i;
            return true;
        }
    }
    return false;
}

function updateAppItems() {
    let displayedItems = [];
    if (totalEntries == 1) {
        displayedItems.push("1 " + Tools.Resource("Item"));
    }
    else {
        displayedItems.push(totalEntries + " " + Tools.Resource("Items"));
    }

    const selectionCount = Object.keys(selection).length;
    if (selectionCount > 0) {
        displayedItems.push(selectionCount + " " + Tools.Resource("Selected"));

        const selectedItem = document.getElementById("e#" + selectionIndex);
        if (selectedItem) {
            let selected = selectedItem.parsingName;
            if (selectionCount > 1) {
                selected += " ...";
            }
            displayedItems.push(selected);
        }
    }

    const items = document.getElementById("app-items");
    items.innerHTML = displayedItems.join(", ");
}

function filterEntries() {
    updateEntries();
}

function updateEntries(folder, parsingName, enumerateOptions) {
    entriesParent.innerHTML = "";

    updateAppItems();

    folder = folder || window.folder;

    switch (window.viewBy) {
        case Enums.ViewBy.Details:
            totalEntries = updateEntriesDetails(folder, enumerateOptions);
            break;

        case Enums.ViewBy.Images:
            totalEntries = updateEntriesImages(folder, enumerateOptions);
            break;
    }

    selectParsingNameEntry(parsingName);
    updateAppItems();
    Window.Window.moveSnappedWindows(); // in case of scrollbars changes
}

function updateEntriesDetails(folder, enumerateOptions) {
    const entries = document.createElement("table");
    entriesParent.appendChild(entries);
    entries.id = "app-entries";
    entries.onclick = e => entryClicked(e);
    entries.ondblclick = e => entryDoubleClicked(e);

    const iconUrl = window.serverUrl + "Icon/";

    let imageOptions = "";
    if (window.viewByDetailsOptions & Enums.ViewByDetailsOptions.ShowIcons) {
        imageOptions = "?options=8"; // GetImageOptions.IconOnly
    }

    let total = 0;
    let startIndex = 0;
    const pageSize = 0;

    do {
        const children = folder.getChildrenView(
            Enums.EntryViewType.Details,
            enumerateOptions || window.entryEnumerateOptions,
            window.sortBy,
            window.sortDirection,
            viewFilterInput.value,
            startIndex,
            pageSize
        );

        const viewInfo = children[0];

        for (let i = 1; i < children.length; i++) {
            const view = children[i].split("|");
            const row = entries.insertRow();
            row.id = "e#" + total;
            const parsingName = view[0];
            const options = view[4] || Enums.WebEntryOptions.None;
            const isFolder = options & Enums.WebEntryOptions.IsFolder;
            const isHidden = options & Enums.WebEntryOptions.IsHidden;
            const isSystem = options & Enums.WebEntryOptions.IsSystem;
            const isCut = options & Enums.WebEntryOptions.IsCut;
            const isCompressed = options & Enums.WebEntryOptions.IsCompressed;
            row.parsingName = parsingName;
            row.isFolder = isFolder;
            row.isCompressed = isCompressed;

            if (selection[total]) {
                row.classList.add("selected");
            }

            if (isFolder) {
                row.classList.add("folder");
            }

            if (isHidden) {
                row.classList.add("hidden");
            }

            if (isCut) {
                row.classList.add("cut");
            }

            if (isSystem) {
                row.classList.add("system");
            }

            if (isCompressed) {
                row.classList.add("compressed");
            }

            const iconCell = row.insertCell();
            const img = new Image();
            img.src = iconUrl + encodeURIComponent(parsingName) + imageOptions;
            iconCell.appendChild(img);

            const dateCell = row.insertCell();
            const modDate = view[1];
            dateCell.innerText = modDate;

            const sizeCell = row.insertCell();
            const size = view[2];
            sizeCell.innerText = size;

            const nameCell = row.insertCell();
            nameCell.innerText = view[3];

            total++;

            if (total % 1000 == 0) {
                dotnet.showLoading(true, "web");
            }
        }

        if (viewInfo.isLast)
            break;

        startIndex += pageSize;
    }
    while (true);

    // make sure there's one more row to avoid scroll issues
    entries.insertRow().insertCell().innerHTML = "&nbsp;";

    dotnet.showLoading(false);
    return total;
}

function updateEntriesImages(folder, enumerateOptions) {
    const children = folder.getChildrenView(
        Enums.EntryViewType.Images,
        enumerateOptions || window.entryEnumerateOptions,
        window.sortBy,
        window.sortDirection,
        viewFilterInput.value
    );

    const entries = document.createElement("div");
    entries.onclick = e => entryClicked(e);
    entries.ondblclick = e => entryDoubleClicked(e);
    entriesParent.appendChild(entries);
    entries.id = "app-images";

    const size = window.thumbnailsSize;
    const iconUrl = window.serverUrl + "Icon/";
    const imageUrl = window.serverUrl + "Image/";
    const pdfPageUrl = window.serverUrl + "PdfPage/";

    const displayTitle = window.viewByImageOptions & Enums.ViewByImageOptions.DisplayTitle;
    const renderPdfThumbnails = window.viewByImageOptions & Enums.ViewByImageOptions.RenderPdfThumbnails;
    const squareThumbnails = window.viewByImageOptions & Enums.ViewByImageOptions.SquareThumbnails;

    let total = 0;
    for (let i = 1; i < children.length; i++) {
        const view = children[i].split("|");
        let parsingName = view[0];
        const options = view[4] || Enums.WebEntryOptions.None;
        const isFolder = options & Enums.WebEntryOptions.IsFolder;
        const isHidden = options & Enums.WebEntryOptions.IsHidden;
        const isSystem = options & Enums.WebEntryOptions.IsSystem;
        const isCompressed = options & Enums.WebEntryOptions.IsCompressed;
        const isPdf = options & Enums.WebEntryOptions.IsPdf;
        const isCut = options & Enums.WebEntryOptions.IsCut;
        const notNative = options & Enums.WebEntryOptions.IsNotWebView2NativeImage;

        const container = document.createElement("div");
        container.id = "e#" + total;
        container.isFolder = isFolder;
        container.isCompressed = isCompressed;
        container.parsingName = parsingName;
        container.className = "app-image-container";
        container.style.width = size + "px";

        // object is selected
        if (selection[total]) {
            container.classList.add("selected");
        }

        let tt;
        if (isFolder) {
            tt = "<center><b>" + view[3] + "</b><br/>" + view[1] + "</center>";
            container.classList.add("folder");
        }
        else {
            tt = "<center><b>" + view[3] + "</b><br/>" + view[2] + "<br/>" + view[1] + "</center>";
        }
        container.setAttribute("tooltip", tt);

        if (isHidden) {
            container.classList.add("hidden");
        }

        if (isSystem) {
            container.classList.add("system");
        }

        if (isCut) {
            container.classList.add("cut");
        }

        if (isCompressed) {
            container.classList.add("compressed");
        }

        const img = new Image();
        if (!isFolder) {
            img.onerror = container.remove();
        }

        if (squareThumbnails) {
            img.style.objectFit = "cover";
            img.style.width = size + "px";
            img.style.height = size + "px";
        }
        else {
            img.style.maxWidth = size + "px";
        }

        if (isFolder) {
            img.src = imageUrl + encodeURIComponent(parsingName) + "?size=" + size;
        }
        else if (notNative) {
            // image mode: only images
            if (window.viewBy == Enums.ViewBy.Images) {
                if (isPdf && renderPdfThumbnails) {
                    img.src = pdfPageUrl + encodeURIComponent(parsingName) + "?size=" + size;
                }
                else {
                    img.src = imageUrl + encodeURIComponent(parsingName) + "?size=" + size;
                }
            }
            else {
                // preview mode: images or previews
                img.src = imageUrl + encodeURIComponent(parsingName) + "?size=" + size;
            }
        }
        else {
            img.src = Tools.encodeFilePathForUrl(parsingName);
        }

        container.appendChild(img);

        if (displayTitle) {
            const titleDiv = document.createElement("div");
            titleDiv.className = "app-image-title";
            titleDiv.innerText = view[3];
            container.appendChild(titleDiv);
        }
        entries.appendChild(container);

        total++;
    }
    return total;
}

#include "../../src/PluginInterface.h"
#include "picojson.h"
#include "resource.h"
#include <windows.h>
#include <commctrl.h>
#include <shlwapi.h>
#include <xmllite.h>
#include <atlbase.h>
#include <string>
#include <vector>
#include <algorithm>
#include <memory>
#include <sstream>
#include <iomanip>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "xmllite.lib")

HINSTANCE g_hModule = NULL;

enum class NodeType { Object, Array, String, Number, Boolean, Null, XmlElement, XmlAttribute, XmlText };

struct DataNode : std::enable_shared_from_this<DataNode> {
    std::wstring name;
    std::wstring value;
    NodeType type;
    std::vector<std::shared_ptr<DataNode>> children;
    std::weak_ptr<DataNode> parent;
    
    DataNode(const std::wstring& n, const std::wstring& v, NodeType t) : name(n), value(v), type(t) {}
    
    void AddChild(std::shared_ptr<DataNode> child) {
        children.push_back(child);
    }
};

std::shared_ptr<DataNode> g_RootNode;
DataNode* g_SelectedNode = nullptr;
std::wstring g_CurrentFormat;
std::wstring g_FilterText;

std::string WideToUtf8(const std::wstring& wstr) {
    if (wstr.empty()) return std::string();
    int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), NULL, 0, NULL, NULL);
    std::string strTo(size_needed, 0);
    WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
    return strTo;
}

std::wstring Utf8ToWide(const std::string& str) {
    if (str.empty()) return std::wstring();
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), NULL, 0);
    std::wstring wstrTo(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
    return wstrTo;
}

// --- JSON Parsing ---
std::shared_ptr<DataNode> ParseJsonValue(const std::wstring& name, const picojson::value& v) {
    if (v.is<picojson::object>()) {
        auto node = std::make_shared<DataNode>(name, L"", NodeType::Object);
        const auto& obj = v.get<picojson::object>();
        for (const auto& kv : obj) {
            auto child = ParseJsonValue(Utf8ToWide(kv.first), kv.second);
            child->parent = node;
            node->AddChild(child);
        }
        return node;
    } else if (v.is<picojson::array>()) {
        auto node = std::make_shared<DataNode>(name, L"", NodeType::Array);
        const auto& arr = v.get<picojson::array>();
        for (size_t i = 0; i < arr.size(); ++i) {
            auto child = ParseJsonValue(L"[" + std::to_wstring(i) + L"]", arr[i]);
            child->parent = node;
            node->AddChild(child);
        }
        return node;
    } else if (v.is<std::string>()) {
        return std::make_shared<DataNode>(name, Utf8ToWide(v.get<std::string>()), NodeType::String);
    } else if (v.is<double>()) {
        return std::make_shared<DataNode>(name, std::to_wstring(v.get<double>()), NodeType::Number);
    } else if (v.is<bool>()) {
        return std::make_shared<DataNode>(name, v.get<bool>() ? L"true" : L"false", NodeType::Boolean);
    } else {
        return std::make_shared<DataNode>(name, L"null", NodeType::Null);
    }
}

// --- XML Parsing ---
std::shared_ptr<DataNode> ParseXml(const std::wstring& xml) {
    CComPtr<IStream> pStream;
    std::string xmlUtf8 = WideToUtf8(xml);
    pStream.Attach(SHCreateMemStream((const BYTE*)xmlUtf8.c_str(), (UINT)xmlUtf8.size()));
    
    CComPtr<IXmlReader> pReader;
    CreateXmlReader(__uuidof(IXmlReader), (void**)&pReader, NULL);
    pReader->SetInput(pStream);
    
    std::shared_ptr<DataNode> root = std::make_shared<DataNode>(L"XML", L"", NodeType::XmlElement);
    std::shared_ptr<DataNode> current = root;
    
    XmlNodeType nodeType;
    const WCHAR* localName;
    const WCHAR* value;
    
    while (S_OK == pReader->Read(&nodeType)) {
        switch (nodeType) {
            case XmlNodeType_Element: {
                pReader->GetLocalName(&localName, NULL);
                auto node = std::make_shared<DataNode>(localName, L"", NodeType::XmlElement);
                node->parent = current;
                current->AddChild(node);
                current = node;
                
                // Attributes
                if (S_OK == pReader->MoveToFirstAttribute()) {
                    do {
                        pReader->GetLocalName(&localName, NULL);
                        pReader->GetValue(&value, NULL);
                        auto attr = std::make_shared<DataNode>(localName, value, NodeType::XmlAttribute);
                        attr->parent = current;
                        current->AddChild(attr);
                    } while (S_OK == pReader->MoveToNextAttribute());
                }
                break;
            }
            case XmlNodeType_EndElement:
                if (!current->parent.expired()) {
                    current = current->parent.lock();
                }
                break;
            case XmlNodeType_Text:
            case XmlNodeType_CDATA:
                pReader->GetValue(&value, NULL);
                auto text = std::make_shared<DataNode>(L"#text", value, NodeType::XmlText);
                text->parent = current;
                current->AddChild(text);
                break;
        }
    }
    // The root we created is a dummy container, return its first child if exists
    if (!root->children.empty()) return root->children[0];
    return root;
}

// --- CSV Parsing ---
std::shared_ptr<DataNode> ParseCsv(const std::wstring& csv) {
    auto root = std::make_shared<DataNode>(L"CSV", L"", NodeType::Array);
    std::wstringstream ss(csv);
    std::wstring line;
    int rowIdx = 0;
    
    while (std::getline(ss, line)) {
        if (line.empty()) continue;
        if (line.back() == L'\r') line.pop_back();
        
        auto rowNode = std::make_shared<DataNode>(L"Row " + std::to_wstring(rowIdx++), L"", NodeType::Array);
        rowNode->parent = root;
        root->AddChild(rowNode);
        
        std::wstringstream lineStream(line);
        std::wstring cell;
        int colIdx = 0;
        
        // Simple CSV split (doesn't handle quotes perfectly but good enough for now)
        while (std::getline(lineStream, cell, L',')) {
            auto cellNode = std::make_shared<DataNode>(L"Col " + std::to_wstring(colIdx++), cell, NodeType::String);
            cellNode->parent = rowNode;
            rowNode->AddChild(cellNode);
        }
    }
    return root;
}

// --- Serialization ---
void SerializeJsonRecursive(std::shared_ptr<DataNode> node, std::wstring& out, int indent) {
    std::wstring indStr(indent * 2, L' ');
    
    if (node->type == NodeType::Object) {
        out += L"{\r\n";
        for (size_t i = 0; i < node->children.size(); ++i) {
            out += indStr + L"  \"" + node->children[i]->name + L"\": ";
            SerializeJsonRecursive(node->children[i], out, indent + 1);
            if (i < node->children.size() - 1) out += L",";
            out += L"\r\n";
        }
        out += indStr + L"}";
    } else if (node->type == NodeType::Array) {
        out += L"[\r\n";
        for (size_t i = 0; i < node->children.size(); ++i) {
            out += indStr + L"  ";
            SerializeJsonRecursive(node->children[i], out, indent + 1);
            if (i < node->children.size() - 1) out += L",";
            out += L"\r\n";
        }
        out += indStr + L"]";
    } else if (node->type == NodeType::String) {
        out += L"\"" + node->value + L"\"";
    } else if (node->type == NodeType::Number || node->type == NodeType::Boolean) {
        out += node->value;
    } else {
        out += L"null";
    }
}

void SerializeXmlRecursive(std::shared_ptr<DataNode> node, std::wstring& out, int indent) {
    std::wstring indStr(indent * 2, L' ');
    
    if (node->type == NodeType::XmlElement) {
        out += indStr + L"<" + node->name;
        
        // Attributes
        for (auto& child : node->children) {
            if (child->type == NodeType::XmlAttribute) {
                out += L" " + child->name + L"=\"" + child->value + L"\"";
            }
        }
        
        bool hasContent = false;
        for (auto& child : node->children) {
            if (child->type != NodeType::XmlAttribute) {
                hasContent = true;
                break;
            }
        }
        
        if (hasContent) {
            out += L">\r\n";
            for (auto& child : node->children) {
                if (child->type != NodeType::XmlAttribute) {
                    SerializeXmlRecursive(child, out, indent + 1);
                }
            }
            out += indStr + L"</" + node->name + L">\r\n";
        } else {
            out += L" />\r\n";
        }
    } else if (node->type == NodeType::XmlText) {
        out += indStr + node->value + L"\r\n";
    }
}

void SerializeCsvRecursive(std::shared_ptr<DataNode> node, std::wstring& out) {
    if (node->type == NodeType::Array) { // Root or Row
        for (auto& child : node->children) {
            if (child->type == NodeType::Array) { // Row
                SerializeCsvRecursive(child, out);
                out += L"\r\n";
            } else { // Cell
                out += child->value + L",";
            }
        }
        if (!node->children.empty() && node->children[0]->type != NodeType::Array && !out.empty() && out.back() == L',') {
             out.pop_back(); // Remove trailing comma for row
        }
    }
}

// --- UI Logic ---
bool ContainsFilter(const std::wstring& text) {
    if (g_FilterText.empty()) return true;
    auto it = std::search(
        text.begin(), text.end(),
        g_FilterText.begin(), g_FilterText.end(),
        [](wchar_t c1, wchar_t c2) {
            return towlower(c1) == towlower(c2);
        }
    );
    return it != text.end();
}

void PopulateTree(HWND hTree, HTREEITEM hParent, std::shared_ptr<DataNode> node) {
    if (!node) return;
    
    std::wstring display = node->name;
    if (node->type != NodeType::Object && node->type != NodeType::Array && node->type != NodeType::XmlElement) {
        display += L": " + node->value;
    }
    
    bool match = ContainsFilter(display);
    bool hasMatchingChildren = false;
    
    // Check children first if we are filtering
    if (!g_FilterText.empty()) {
        // This is a simple check, for full filtering we need to be smarter
        // But for now, let's just add everything and rely on user expanding
        // Or we can implement the same logic as before
    }

    TVINSERTSTRUCT tvis = {0};
    tvis.hParent = hParent;
    tvis.hInsertAfter = TVI_LAST;
    tvis.item.mask = TVIF_TEXT | TVIF_PARAM;
    tvis.item.pszText = const_cast<LPWSTR>(display.c_str());
    tvis.item.lParam = (LPARAM)node.get();
    
    HTREEITEM hItem = TreeView_InsertItem(hTree, &tvis);
    
    for (auto& child : node->children) {
        PopulateTree(hTree, hItem, child);
    }
    
    if (!g_FilterText.empty()) {
        TreeView_Expand(hTree, hItem, TVE_EXPAND);
    }
}

void UpdateRightPanel(HWND hDlg, DataNode* node) {
    if (!node) return;
    
    SetDlgItemText(hDlg, IDC_EDIT_NAME, node->name.c_str());
    SetDlgItemText(hDlg, IDC_EDIT_VALUE, node->value.c_str());
    
    HWND hCombo = GetDlgItem(hDlg, IDC_COMBO_TYPE);
    SendMessage(hCombo, CB_RESETCONTENT, 0, 0);
    
    const wchar_t* types[] = { L"Object", L"Array", L"String", L"Number", L"Boolean", L"Null", L"Element", L"Attribute", L"Text" };
    for (auto t : types) SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)t);
    
    SendMessage(hCombo, CB_SETCURSEL, (WPARAM)node->type, 0);
}

void SaveNodeFromPanel(HWND hDlg, DataNode* node) {
    if (!node) return;
    
    wchar_t buf[1024];
    GetDlgItemText(hDlg, IDC_EDIT_NAME, buf, 1024);
    node->name = buf;
    
    // Get value length first
    int len = GetWindowTextLength(GetDlgItem(hDlg, IDC_EDIT_VALUE));
    std::vector<wchar_t> valBuf(len + 1);
    GetDlgItemText(hDlg, IDC_EDIT_VALUE, valBuf.data(), len + 1);
    node->value = valBuf.data();
    
    int typeIdx = SendMessage(GetDlgItem(hDlg, IDC_COMBO_TYPE), CB_GETCURSEL, 0, 0);
    if (typeIdx >= 0) node->type = (NodeType)typeIdx;
    
    // Update tree item text
    // This requires finding the tree item corresponding to this node.
    // We can traverse the tree or just refresh the whole tree. Refreshing is easier.
}

INT_PTR CALLBACK DataViewerDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        {
            HWND hTree = GetDlgItem(hDlg, IDC_DATA_TREE);
            PopulateTree(hTree, TVI_ROOT, g_RootNode);
        }
        return (INT_PTR)TRUE;

    case WM_NOTIFY:
        {
            LPNMHDR lpnmh = (LPNMHDR)lParam;
            if (lpnmh->idFrom == IDC_DATA_TREE && lpnmh->code == TVN_SELCHANGED) {
                LPNMTREEVIEW pnmtv = (LPNMTREEVIEW)lParam;
                g_SelectedNode = (DataNode*)pnmtv->itemNew.lParam;
                UpdateRightPanel(hDlg, g_SelectedNode);
            }
        }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        if (LOWORD(wParam) == IDC_BTN_UPDATE) {
            if (g_SelectedNode) {
                SaveNodeFromPanel(hDlg, g_SelectedNode);
                // Refresh tree text for selected item
                HWND hTree = GetDlgItem(hDlg, IDC_DATA_TREE);
                HTREEITEM hItem = TreeView_GetSelection(hTree);
                if (hItem) {
                    std::wstring display = g_SelectedNode->name;
                    if (g_SelectedNode->type != NodeType::Object && g_SelectedNode->type != NodeType::Array && g_SelectedNode->type != NodeType::XmlElement) {
                        display += L": " + g_SelectedNode->value;
                    }
                    TVITEM tvi = {0};
                    tvi.hItem = hItem;
                    tvi.mask = TVIF_TEXT;
                    tvi.pszText = const_cast<LPWSTR>(display.c_str());
                    TreeView_SetItem(hTree, &tvi);
                }
            }
        }
        if (LOWORD(wParam) == IDC_BTN_ADD) {
            if (g_SelectedNode) {
                auto child = std::make_shared<DataNode>(L"New Node", L"", NodeType::String);
                child->parent = g_SelectedNode->shared_from_this();
                g_SelectedNode->AddChild(child);
                
                // Add to tree
                HWND hTree = GetDlgItem(hDlg, IDC_DATA_TREE);
                HTREEITEM hItem = TreeView_GetSelection(hTree);
                PopulateTree(hTree, hItem, child);
                TreeView_Expand(hTree, hItem, TVE_EXPAND);
            }
        }
        if (LOWORD(wParam) == IDC_BTN_DELETE) {
            if (g_SelectedNode && !g_SelectedNode->parent.expired()) {
                auto parent = g_SelectedNode->parent.lock();
                auto it = std::find_if(parent->children.begin(), parent->children.end(), 
                    [](const std::shared_ptr<DataNode>& n) { return n.get() == g_SelectedNode; });
                if (it != parent->children.end()) {
                    parent->children.erase(it);
                    
                    HWND hTree = GetDlgItem(hDlg, IDC_DATA_TREE);
                    HTREEITEM hItem = TreeView_GetSelection(hTree);
                    TreeView_DeleteItem(hTree, hItem);
                    g_SelectedNode = nullptr;
                    
                    // Clear right panel
                    SetDlgItemText(hDlg, IDC_EDIT_NAME, L"");
                    SetDlgItemText(hDlg, IDC_EDIT_VALUE, L"");
                }
            }
        }
        if (LOWORD(wParam) == IDC_BTN_SAVE) {
            std::wstring out;
            if (g_CurrentFormat == L"JSON") SerializeJsonRecursive(g_RootNode, out, 0);
            else if (g_CurrentFormat == L"XML") SerializeXmlRecursive(g_RootNode, out, 0);
            else if (g_CurrentFormat == L"CSV") SerializeCsvRecursive(g_RootNode, out);
            
            HWND hEditor = GetParent(hDlg);
            SetWindowText(hEditor, out.c_str());
            EndDialog(hDlg, IDOK);
        }
        break;
    }
    return (INT_PTR)FALSE;
}

extern "C" {
    PLUGIN_API const wchar_t* GetPluginName() { return L"Data Viewer"; }
    PLUGIN_API const wchar_t* GetPluginDescription() { return L"View and Edit JSON/XML/CSV files."; }
    PLUGIN_API const wchar_t* GetPluginVersion() { return L"2.0"; }

    PLUGIN_API const wchar_t* GetPluginStatus(const wchar_t* filePath) {
        if (filePath == nullptr) return nullptr;
        const wchar_t* ext = PathFindExtension(filePath);
        if (_wcsicmp(ext, L".json") == 0) return L"JSON";
        if (_wcsicmp(ext, L".xml") == 0) return L"XML";
        if (_wcsicmp(ext, L".csv") == 0) return L"CSV";
        return nullptr;
    }

    void ViewData(HWND hEditorWnd) {
        int len = GetWindowTextLength(hEditorWnd);
        if (len <= 0) return;

        std::vector<wchar_t> buffer(len + 1);
        GetWindowText(hEditorWnd, buffer.data(), len + 1);
        std::wstring text(buffer.data());
        
        // Determine format based on content or file extension?
        // We don't have file extension here easily unless we store it.
        // But we can guess.
        if (text.find(L"{") == 0 || text.find(L"[") == 0) g_CurrentFormat = L"JSON";
        else if (text.find(L"<") == 0) g_CurrentFormat = L"XML";
        else g_CurrentFormat = L"CSV"; // Fallback
        
        if (g_CurrentFormat == L"JSON") {
            std::string jsonUtf8 = WideToUtf8(text);
            picojson::value v;
            std::string err;
            picojson::parse(v, jsonUtf8.begin(), jsonUtf8.end(), &err);
            if (err.empty()) g_RootNode = ParseJsonValue(L"Root", v);
            else { MessageBox(hEditorWnd, Utf8ToWide(err).c_str(), L"JSON Error", MB_OK); return; }
        } else if (g_CurrentFormat == L"XML") {
            CoInitialize(NULL);
            g_RootNode = ParseXml(text);
        } else {
            g_RootNode = ParseCsv(text);
        }

        DialogBox(g_hModule, MAKEINTRESOURCE(IDD_DATA_VIEWER), hEditorWnd, DataViewerDlgProc);
        
        if (g_CurrentFormat == L"XML") CoUninitialize();
        g_RootNode.reset();
    }

    PluginMenuItem g_Items[] = {
        { L"View Data Structure", ViewData, L"Ctrl+Shift+V" }
    };

    PLUGIN_API PluginMenuItem* GetPluginMenuItems(int* count) {
        *count = sizeof(g_Items) / sizeof(g_Items[0]);
        return g_Items;
    }
}

BOOL APIENTRY DllMain( HMODULE hModule, DWORD  ul_reason_for_call, LPVOID lpReserved ) {
    if (ul_reason_for_call == DLL_PROCESS_ATTACH) g_hModule = hModule;
    return TRUE;
}

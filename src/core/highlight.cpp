#include "highlight.h"

#include <algorithm>
#include <cstring>
#include <mutex>
#include <string>
#include <unordered_map>
#include <unordered_set>

namespace jnp {

// ---------------------------------------------------------------- language table

enum : u32 {
    F_CASELESS = 1u << 0,      // keywords ignore case (SQL, VB, Pascal, ...)
    F_PREPROC_HASH = 1u << 1,  // "#" starting a line begins a preprocessor line (C family, shaders)
    F_TRIPLE = 1u << 2,        // """ and ''' strings across lines (Python, Kotlin, Swift, Scala, Julia, Dart)
    F_BACKTICK = 1u << 3,      // `...` across lines (JS template strings, Go raw strings)
    F_CPP_RAW = 1u << 4,       // R"( ... )"
    F_CS_VERBATIM = 1u << 5,   // @"..." with "" escapes
    F_LUA = 1u << 6,           // [[ ]] strings and --[[ ]] comments
    F_MARKUP = 1u << 7,        // XML / HTML
    F_MARKDOWN = 1u << 8,
    F_INI = 1u << 9,           // [section], key = value
    F_YAML = 1u << 10,
    F_JSON = 1u << 11,         // "key": marked as attribute
    F_DOLLAR_VARS = 1u << 12,  // $name is a variable (shell, PowerShell, PHP, Perl)
    F_DASH_IDENT = 1u << 13,   // identifiers may contain "-" (CSS, PowerShell cmdlets, CMake)
    F_AT_ANNOT = 1u << 14,     // @Name is an annotation / decorator
    F_MULTILINE_STR = 1u << 15,  // "..." may continue on the next line (shell, PowerShell, Rust, PHP)
    F_LIFETIMES = 1u << 16,    // 'a is a Rust lifetime, not a character
    F_DIFF = 1u << 17,
    F_CSS = 1u << 18,          // property: value; #id, .class
    F_DOLLAR_IDENT = 1u << 19, // "$" may be part of an identifier (JS, Java)
    F_BATCH = 1u << 20,        // REM / :: comments, %var%
    F_ASM = 1u << 21,          // ; comments, labels
};

struct Lang {
    const char* name;
    const char* exts;      // space-separated, lower case; "=name" matches a whole file name
    const char* keywords;
    const char* types;
    const char* literals;
    const char* macros;    // shown as preprocessor (Unreal's UCLASS, ...)
    const char* line1;     // line comment markers
    const char* line2;
    const char* blockOpen; // block comment
    const char* blockClose;
    bool nested;           // block comments nest
    const char* quotes;    // characters that start a one-line string
    u32 flags;
};

#define C_KEYWORDS                                                                                               \
    "auto break case const continue default do else enum extern for goto if inline register restrict return "  \
    "sizeof static struct switch typedef union volatile while _Alignas _Alignof _Atomic _Generic _Noreturn "      \
    "_Static_assert _Thread_local "
#define C_TYPES                                                                                                  \
    "void char short int long float double signed unsigned bool _Bool size_t ssize_t ptrdiff_t intptr_t "         \
    "uintptr_t int8_t int16_t int32_t int64_t uint8_t uint16_t uint32_t uint64_t wchar_t char8_t char16_t "      \
    "char32_t FILE "
#define CPP_KEYWORDS                                                                                             \
    C_KEYWORDS                                                                                                   \
    "alignas alignof and and_eq asm bitand bitor catch class co_await co_return co_yield compl concept "          \
    "consteval constexpr constinit const_cast decltype delete dynamic_cast explicit export final friend import " \
    "module mutable namespace new noexcept not not_eq operator or or_eq override private protected public "     \
    "reinterpret_cast requires static_assert static_cast template this thread_local throw try typeid typename " \
    "using virtual xor xor_eq __declspec __forceinline __cdecl __stdcall __fastcall __int64 FORCEINLINE "
#define UNREAL_TYPES                                                                                              \
    "int8 int16 int32 int64 uint8 uint16 uint32 uint64 TCHAR ANSICHAR WIDECHAR UTF8CHAR FString FName FText "     \
    "TArray TMap TSet TMultiMap TSharedPtr TSharedRef TWeakPtr TUniquePtr TWeakObjectPtr TObjectPtr "             \
    "TSoftObjectPtr TSoftClassPtr TSubclassOf TOptional TVariant TFunction TDelegate TEnumAsByte TScriptInterface " \
    "FVector FVector2D FVector4 FRotator FQuat FTransform FColor FLinearColor FMatrix FBox FIntPoint FIntVector "  \
    "FGuid FDateTime FTimespan UObject AActor APawn ACharacter APlayerController AGameModeBase UActorComponent "   \
    "USceneComponent UWorld UClass UEnum UStruct UFunction FProperty "
#define UNREAL_MACROS                                                                                             \
    "UCLASS USTRUCT UENUM UFUNCTION UPROPERTY UPARAM UMETA UINTERFACE UDELEGATE GENERATED_BODY "                   \
    "GENERATED_UCLASS_BODY GENERATED_USTRUCT_BODY GENERATED_IINTERFACE_BODY DECLARE_DYNAMIC_MULTICAST_DELEGATE "   \
    "DECLARE_DELEGATE DECLARE_MULTICAST_DELEGATE DECLARE_LOG_CATEGORY_EXTERN DEFINE_LOG_CATEGORY UE_LOG check "     \
    "checkf verify verifyf ensure ensureMsgf ensureAlways TEXT LOCTEXT NSLOCTEXT TRACE_CPUPROFILER_EVENT_SCOPE "    \
    "SCOPE_CYCLE_COUNTER QUICK_SCOPE_CYCLE_COUNTER UE_DEPRECATED PRAGMA_DISABLE_DEPRECATION_WARNINGS "            \
    "PRAGMA_ENABLE_DEPRECATION_WARNINGS "
#define C_LITERALS "true false NULL nullptr TRUE FALSE "
#define JS_KEYWORDS                                                                                              \
    "async await break case catch class const continue debugger default delete do else export extends finally " \
    "for from function get if import in instanceof let new of return set static super switch this throw try "   \
    "typeof var void while with yield "
#define SHADER_KEYWORDS                                                                                          \
    "break case cbuffer centroid class const continue default discard do else export extern for groupshared if "  \
    "in inline inout interface linear namespace nointerpolation noperspective out precise return register shared " \
    "static struct switch tbuffer typedef uniform volatile while layout attribute varying buffer coherent "        \
    "readonly writeonly restrict flat smooth invariant subroutine highp mediump lowp precision patch sample "

static const Lang kLangs[] = {
    {"C++", "cpp cc cxx c++ hpp hh hxx h++ inl ipp tpp ixx cppm ino h m mm", CPP_KEYWORDS, C_TYPES UNREAL_TYPES,
     C_LITERALS, UNREAL_MACROS, "//", nullptr, "/*", "*/", false, "\"'", F_PREPROC_HASH | F_CPP_RAW},
    {"C", "c", C_KEYWORDS, C_TYPES, C_LITERALS, nullptr, "//", nullptr, "/*", "*/", false, "\"'", F_PREPROC_HASH},
    {"C#", "cs csx",
     "abstract as base break case catch checked class const continue default delegate do else enum event explicit "
     "extern finally fixed for foreach goto if implicit in interface internal is lock namespace new operator out "
     "override params private protected public readonly ref return sealed sizeof stackalloc static struct switch "
     "this throw try typeof unchecked unsafe using virtual volatile while add alias async await by descending "
     "dynamic equals from get global group init into join let nameof notnull on orderby partial record remove "
     "required select set unmanaged value var when where with yield file scoped",
     "bool byte char decimal double float int long object sbyte short string uint ulong ushort void nint nuint "
     "String Object Task List Dictionary IEnumerable Action Func",
     "true false null", nullptr, "//", nullptr, "/*", "*/", false, "\"'", F_PREPROC_HASH | F_CS_VERBATIM | F_AT_ANNOT},
    {"Java", "java",
     "abstract assert break case catch class const continue default do else enum extends final finally for goto "
     "if implements import instanceof interface native new package private protected public return static "
     "strictfp super switch synchronized this throw throws transient try volatile while var record sealed permits "
     "yield non-sealed",
     "boolean byte char double float int long short void String Object Integer Long Double Boolean List Map Set",
     "true false null", nullptr, "//", nullptr, "/*", "*/", false, "\"'", F_AT_ANNOT | F_DOLLAR_IDENT | F_TRIPLE},
    {"Kotlin", "kt kts",
     "as break class continue do else for fun if in interface is null object package return super this throw try "
     "typealias typeof val var when while by catch constructor delegate dynamic field file finally get import "
     "init param property receiver set setparam where actual abstract annotation companion const crossinline data "
     "enum expect external final infix inline inner internal lateinit noinline open operator out override private "
     "protected public reified sealed suspend tailrec vararg",
     "Any Unit Nothing Int Long Short Byte Double Float Boolean Char String Array List Map Set", "true false null",
     nullptr, "//", nullptr, "/*", "*/", true, "\"'", F_AT_ANNOT | F_TRIPLE},
    {"Swift", "swift",
     "associatedtype class deinit enum extension fileprivate func import init inout internal let open operator "
     "private protocol public rethrows static struct subscript typealias var break case continue default defer do "
     "else fallthrough for guard if in repeat return switch where while as catch is super self Self throw throws "
     "try async await some any actor nonisolated mutating override final lazy weak unowned convenience required",
     "Int Int8 Int16 Int32 Int64 UInt UInt8 UInt16 UInt32 UInt64 Float Double Bool String Character Array "
     "Dictionary Set Optional Void Any AnyObject",
     "true false nil", nullptr, "//", nullptr, "/*", "*/", true, "\"", F_AT_ANNOT | F_TRIPLE | F_PREPROC_HASH},
    {"Go", "go",
     "break case chan const continue default defer else fallthrough for func go goto if import interface map "
     "package range return select struct switch type var",
     "bool byte complex64 complex128 error float32 float64 int int8 int16 int32 int64 rune string uint uint8 "
     "uint16 uint32 uint64 uintptr any comparable",
     "true false nil iota", nullptr, "//", nullptr, "/*", "*/", false, "\"'", F_BACKTICK},
    {"Rust", "rs",
     "as async await break const continue crate dyn else enum extern fn for if impl in let loop match mod move mut "
     "pub ref return self Self static struct super trait type union unsafe use where while yield macro_rules",
     "bool char str u8 u16 u32 u64 u128 usize i8 i16 i32 i64 i128 isize f32 f64 String Vec Option Result Box Rc "
     "Arc HashMap HashSet",
     "true false None Some Ok Err", nullptr, "//", nullptr, "/*", "*/", true, "\"'",
     F_MULTILINE_STR | F_LIFETIMES | F_PREPROC_HASH},
    {"JavaScript", "js mjs cjs jsx", JS_KEYWORDS, "Array Object String Number Boolean Promise Map Set Date RegExp Error Math JSON",
     "true false null undefined NaN Infinity", nullptr, "//", nullptr, "/*", "*/", false, "\"'",
     F_BACKTICK | F_DOLLAR_IDENT},
    {"TypeScript", "ts tsx mts cts",
     JS_KEYWORDS "abstract as asserts declare enum implements infer interface is keyof namespace private protected "
                 "public readonly satisfies type unique override module",
     "any boolean never number object string symbol unknown void bigint Array Object String Number Boolean "
     "Promise Map Set Date Record Partial Readonly",
     "true false null undefined NaN Infinity", nullptr, "//", nullptr, "/*", "*/", false, "\"'",
     F_BACKTICK | F_DOLLAR_IDENT | F_AT_ANNOT},
    {"Python", "py pyw pyi pyx =sconstruct =sconscript",
     "and as assert async await break class continue def del elif else except finally for from global if import "
     "in is lambda nonlocal not or pass raise return try while with yield match case type",
     "int float complex str bytes bytearray bool list tuple dict set frozenset object type range len print "
     "isinstance super self cls",
     "True False None NotImplemented Ellipsis", nullptr, "#", nullptr, nullptr, nullptr, false, "\"'",
     F_TRIPLE | F_AT_ANNOT},
    {"Ruby", "rb rake gemspec ru =rakefile =gemfile",
     "alias and begin break case class def defined? do else elsif end ensure for if in module next not or redo "
     "rescue retry return self super then undef unless until when while yield require require_relative attr_accessor "
     "attr_reader attr_writer private protected public lambda proc",
     "Array Hash String Integer Float Symbol Object", "true false nil", nullptr, "#", nullptr, "=begin", "=end",
     false, "\"'", F_AT_ANNOT},
    {"PHP", "php phtml",
     "abstract and array as break callable case catch class clone const continue declare default do echo else "
     "elseif empty enddeclare endfor endforeach endif endswitch endwhile enum extends final finally fn for foreach "
     "function global goto if implements include include_once instanceof insteadof interface isset list match "
     "namespace new or print private protected public readonly require require_once return static switch throw "
     "trait try unset use var while xor yield",
     "int float bool string void mixed object iterable never self parent", "true false null TRUE FALSE NULL",
     nullptr, "//", "#", "/*", "*/", false, "\"'", F_DOLLAR_VARS | F_MULTILINE_STR},
    {"Lua", "lua luau",
     "and break do else elseif end for function goto if in local not or repeat return then until while",
     "self", "true false nil", nullptr, "--", nullptr, nullptr, nullptr, false, "\"'", F_LUA},
    {"Perl", "pl pm t",
     "my our local sub if elsif else unless while until for foreach do last next redo return use no require "
     "package and or not eq ne lt gt le ge cmp",
     "", "undef", nullptr, "#", nullptr, nullptr, nullptr, false, "\"'", F_DOLLAR_VARS},
    {"Shell", "sh bash zsh ksh fish =.bashrc =.zshrc =.profile =.bash_profile",
     "if then else elif fi case esac for select while until do done in function time return exit break continue "
     "local export readonly declare typeset unset shift source alias set eval exec trap",
     "echo printf read cd pwd test cat grep sed awk", "true false", nullptr, "#", nullptr, nullptr, nullptr, false,
     "\"'", F_DOLLAR_VARS | F_MULTILINE_STR},
    {"PowerShell", "ps1 psm1 psd1",
     "begin break catch class continue data define do dynamicparam else elseif end enum exit filter finally for "
     "foreach from function hidden if in param process return static switch throw trap try until using var while "
     "workflow -and -or -not -xor -eq -ne -gt -ge -lt -le -like -notlike -match -notmatch -contains -notcontains "
     "-in -notin -replace -is -isnot -as -f",
     "string int long bool double hashtable array object switch pscustomobject", "$true $false $null", nullptr,
     "#", nullptr, "<#", "#>", false, "\"'", F_DOLLAR_VARS | F_DASH_IDENT | F_MULTILINE_STR | F_CASELESS},
    {"Batch", "bat cmd",
     "call cd chdir cls copy del dir echo else endlocal errorlevel exist exit for goto if in md mkdir move not nul "
     "pause popd pushd rd rem ren rename rmdir set setlocal shift start title type do defined",
     "", "", nullptr, "::", nullptr, nullptr, nullptr, false, "\"", F_CASELESS | F_BATCH},
    {"SQL", "sql psql mysql",
     "add all alter and any as asc begin between by case check column commit constraint create cross database "
     "default delete desc distinct drop else end exists foreign from full group having if in index inner insert "
     "into is join key left like limit not null on or order outer primary procedure references replace returning "
     "right rollback select set table then top transaction trigger truncate union unique update values view when "
     "where with function returns declare exec execute go over partition merge using",
     "int integer bigint smallint tinyint bit decimal numeric float real double money char varchar nchar nvarchar "
     "text ntext date datetime datetime2 time timestamp boolean blob uuid serial json jsonb",
     "true false null", nullptr, "--", nullptr, "/*", "*/", false, "'\"", F_CASELESS},
    {"JSON", "json jsonc json5 uproject uplugin code-workspace babelrc eslintrc =.prettierrc", "", "",
     "true false null", nullptr, "//", nullptr, "/*", "*/", false, "\"'", F_JSON},
    {"XML",
     "xml xaml csproj vcxproj vbproj fsproj props targets config manifest resx svg plist xsd xsl xslt nuspec "
     "wsdl rss atom kml gpx storyboard xib",
     "", "", "", nullptr, nullptr, nullptr, nullptr, nullptr, false, "\"'", F_MARKUP},
    {"HTML", "html htm xhtml shtml vue svelte cshtml razor", "", "", "", nullptr, nullptr, nullptr, nullptr, nullptr,
     false, "\"'", F_MARKUP | F_CASELESS},
    {"CSS", "css scss sass less",
     "@media @import @font-face @keyframes @supports @page @charset @namespace @mixin @include @extend @if @else "
     "@each @for @function @return @use @forward !important",
     "", "inherit initial unset none auto", nullptr, "//", nullptr, "/*", "*/", false, "\"'",
     F_CSS | F_DASH_IDENT},
    {"YAML", "yml yaml =.clang-format =.clang-tidy", "", "", "true false null yes no on off ~", nullptr, "#",
     nullptr, nullptr, nullptr, false, "\"'", F_YAML},
    {"TOML", "toml =cargo.lock =pipfile", "", "", "true false", nullptr, "#", nullptr, nullptr, nullptr, false,
     "\"'", F_INI | F_TRIPLE},
    {"INI",
     "ini cfg inf conf properties editorconfig gitconfig gitattributes gitmodules desktop reg "
     "=.editorconfig =.gitconfig =.gitattributes =.npmrc",
     "", "", "true false yes no on off", nullptr, ";", "#", nullptr, nullptr, false, "\"", F_INI},
    {"Markdown", "md markdown mdx", "", "", "", nullptr, nullptr, nullptr, nullptr, nullptr, false, "",
     F_MARKDOWN},
    {"HLSL", "hlsl hlsli fx fxh usf ush shader cginc compute",
     SHADER_KEYWORDS "SV_Position SV_Target SV_Depth SV_DispatchThreadID SV_GroupID SV_GroupThreadID "
                     "SV_VertexID SV_InstanceID SV_IsFrontFace",
     "void bool int uint dword half float double min16float min10float min16int min12int min16uint float2 float3 "
     "float4 float2x2 float3x3 float4x4 half2 half3 half4 int2 int3 int4 uint2 uint3 uint4 bool2 bool3 bool4 "
     "matrix vector Texture1D Texture2D Texture3D TextureCube Texture2DArray RWTexture2D RWTexture3D Buffer "
     "RWBuffer StructuredBuffer RWStructuredBuffer ByteAddressBuffer RWByteAddressBuffer SamplerState "
     "SamplerComparisonState ConstantBuffer FMaterialPixelParameters FMaterialVertexParameters",
     "true false", nullptr, "//", nullptr, "/*", "*/", false, "\"", F_PREPROC_HASH},
    {"GLSL", "glsl vert frag geom comp tesc tese vs fs gs rgen rchit rmiss mesh task",
     SHADER_KEYWORDS,
     "void bool int uint float double vec2 vec3 vec4 dvec2 dvec3 dvec4 ivec2 ivec3 ivec4 uvec2 uvec3 uvec4 bvec2 "
     "bvec3 bvec4 mat2 mat3 mat4 mat2x2 mat3x3 mat4x4 sampler1D sampler2D sampler3D samplerCube sampler2DArray "
     "sampler2DShadow image2D image3D texture2D",
     "true false", nullptr, "//", nullptr, "/*", "*/", false, "\"", F_PREPROC_HASH},
    {"CMake", "cmake =cmakelists.txt",
     "if elseif else endif foreach endforeach while endwhile function endfunction macro endmacro return break "
     "continue set unset option project cmake_minimum_required add_executable add_library add_subdirectory "
     "target_link_libraries target_include_directories target_compile_definitions target_compile_options "
     "include find_package message install list string file configure_file set_target_properties",
     "", "on off true false yes no", nullptr, "#", nullptr, nullptr, nullptr, false, "\"",
     F_CASELESS | F_DOLLAR_VARS},
    {"Makefile", "mk mak =makefile =gnumakefile", "ifeq ifneq ifdef ifndef else endif include define endef export "
                                                   "override vpath .PHONY",
     "", "", nullptr, "#", nullptr, nullptr, nullptr, false, "\"'", F_DOLLAR_VARS},
    {"Dockerfile", "dockerfile =dockerfile =containerfile",
     "from run cmd label maintainer expose env add copy entrypoint volume user workdir arg onbuild stopsignal "
     "healthcheck shell as",
     "", "", nullptr, "#", nullptr, nullptr, nullptr, false, "\"'", F_CASELESS | F_DOLLAR_VARS},
    {"Diff", "diff patch rej", "", "", "", nullptr, nullptr, nullptr, nullptr, nullptr, false, "", F_DIFF},
    {"Dart", "dart",
     "abstract as assert async await base break case catch class const continue covariant default deferred do "
     "dynamic else enum export extends extension external factory final finally for get hide if implements import "
     "in interface is late library mixin new on operator part required rethrow return sealed set show static "
     "super switch sync this throw try typedef var when while with yield",
     "int double num bool String List Map Set Future Stream void Object", "true false null", nullptr, "//",
     nullptr, "/*", "*/", true, "\"'", F_TRIPLE | F_AT_ANNOT | F_DOLLAR_VARS},
    {"Scala", "scala sc sbt",
     "abstract case catch class def do else enum export extends final finally for forSome given if implicit "
     "import lazy match new object override package private protected return sealed super then this throw trait "
     "try type using val var while with yield",
     "Int Long Double Float Boolean Char String Unit Any AnyRef Nothing List Map Set Option Seq",
     "true false null None Some", nullptr, "//", nullptr, "/*", "*/", true, "\"'", F_TRIPLE | F_AT_ANNOT},
    {"F#", "fs fsi fsx",
     "abstract and as assert base begin class default delegate do done downcast downto elif else end exception "
     "extern for fun function global if in inherit inline interface internal lazy let match member module mutable "
     "namespace new of open or override private public rec return static struct then to try type upcast use val "
     "void when while with yield",
     "int float string bool char unit list option seq array", "true false null None Some", nullptr, "//", nullptr,
     "(*", "*)", true, "\"", F_TRIPLE},
    {"Visual Basic", "vb vbs bas cls frm",
     "addhandler addressof alias and andalso as byref byval call case catch class const continue declare default "
     "delegate dim do each else elseif end enum erase error event exit finally for friend function get global "
     "gosub goto handles if implements imports in inherits interface is isnot let lib like loop me mod module "
     "mustinherit mustoverride mybase namespace narrowing new next not nothing notinheritable notoverridable of on "
     "operator option optional or orelse overloads overridable overrides paramarray partial private property "
     "protected public raiseevent readonly redim rem removehandler resume return select set shadows shared static "
     "step stop structure sub synclock then throw to try typeof using when while with withevents writeonly xor",
     "boolean byte char date decimal double integer long object sbyte short single string uinteger ulong ushort "
     "variant",
     "true false nothing", nullptr, "'", nullptr, nullptr, nullptr, false, "\"", F_CASELESS | F_PREPROC_HASH},
    {"Pascal", "pas pp dpr dpk lpr inc",
     "and array as asm begin case class const constructor destructor div do downto else end except exports file "
     "finalization finally for function goto if implementation in inherited initialization inline interface is "
     "label library mod nil not object of on or out packed procedure program property raise record repeat "
     "resourcestring set shl shr string then threadvar to try type unit until uses var while with xor private "
     "protected public published override virtual abstract",
     "integer cardinal shortint smallint longint int64 byte word longword char boolean real single double "
     "extended currency pointer",
     "true false nil", nullptr, "//", nullptr, "{", "}", false, "'", F_CASELESS},
    {"Assembly", "asm s nasm masm inc", "", "", "", nullptr, ";", nullptr, "/*", "*/", false, "\"'",
     F_CASELESS | F_ASM},
    {"GraphQL", "graphql gql",
     "query mutation subscription fragment on type interface union enum input scalar schema extend implements "
     "directive",
     "Int Float String Boolean ID", "true false null", nullptr, "#", nullptr, nullptr, nullptr, false, "\"",
     F_TRIPLE | F_AT_ANNOT | F_DOLLAR_VARS},
    {"Protobuf", "proto",
     "syntax package import option message enum service rpc returns repeated optional required oneof map reserved "
     "extend stream public weak",
     "double float int32 int64 uint32 uint64 sint32 sint64 fixed32 fixed64 sfixed32 sfixed64 bool string bytes",
     "true false", nullptr, "//", nullptr, "/*", "*/", false, "\"'", 0},
    {"Terraform", "tf tfvars hcl nomad",
     "resource data variable output locals module provider terraform for_each count depends_on lifecycle dynamic "
     "for in if",
     "string number bool list map set object tuple any", "true false null", nullptr, "#", "//", "/*", "*/", false,
     "\"", F_DOLLAR_VARS},
    {"Groovy", "groovy gradle gvy",
     "abstract as assert break case catch class const continue def default do else enum extends final finally for "
     "goto if implements import in instanceof interface native new package private protected public return "
     "static super switch synchronized this throw throws trait transient try var volatile while",
     "boolean byte char double float int long short void String Object List Map", "true false null", nullptr, "//",
     nullptr, "/*", "*/", false, "\"'", F_AT_ANNOT | F_TRIPLE | F_DOLLAR_VARS},
    {"Julia", "jl",
     "baremodule begin break catch const continue do else elseif end export finally for function global if import "
     "let local macro module quote return struct try using while mutable abstract primitive type where in isa",
     "Int Int8 Int16 Int32 Int64 UInt8 Float32 Float64 Bool String Char Vector Matrix Array Dict Any Nothing",
     "true false nothing missing NaN Inf", nullptr, "#", nullptr, "#=", "=#", true, "\"'", F_TRIPLE | F_AT_ANNOT},
    {"R", "r rmd",
     "if else repeat while function for in next break library require return switch", "",
     "TRUE FALSE NULL NA NaN Inf T F", nullptr, "#", nullptr, nullptr, nullptr, false, "\"'", 0},
    {"Zig", "zig zon",
     "addrspace align allowzero and anyframe anytype asm async await break callconv catch comptime const continue "
     "defer else enum errdefer error export extern fn for if inline linksection noalias noinline nosuspend opaque "
     "or orelse packed pub resume return struct suspend switch test threadlocal try union unreachable usingnamespace "
     "var volatile while",
     "i8 u8 i16 u16 i32 u32 i64 u64 i128 u128 isize usize f16 f32 f64 f128 bool void anyerror type noreturn",
     "true false null undefined", nullptr, "//", nullptr, nullptr, nullptr, false, "\"'", 0},
    {"Elixir", "ex exs heex",
     "after alias and case catch cond def defdelegate defexception defguard defimpl defmacro defmodule defp "
     "defprotocol defstruct do else end fn for if import in not or quote raise receive require rescue try unless "
     "unquote use when with",
     "", "true false nil", nullptr, "#", nullptr, nullptr, nullptr, false, "\"'", F_TRIPLE | F_AT_ANNOT},
    {"Haskell", "hs lhs",
     "case class data default deriving do else foreign if import in infix infixl infixr instance let module newtype "
     "of then type where qualified as hiding",
     "Int Integer Float Double Bool Char String Maybe Either IO", "True False Nothing Just Left Right", nullptr,
     "--", nullptr, "{-", "-}", true, "\"", 0},
};

// ---------------------------------------------------------------- keyword sets

namespace {

struct Words {
    std::unordered_set<std::string> keywords, types, literals, macros;
};

void addWords(std::unordered_set<std::string>& set, const char* list, bool lower) {
    if (!list) return;
    const char* p = list;
    while (*p) {
        while (*p == ' ') p++;
        const char* s = p;
        while (*p && *p != ' ') p++;
        if (p > s) {
            std::string w(s, (size_t)(p - s));
            if (lower)
                for (char& c : w) c = (char)((c >= 'A' && c <= 'Z') ? c + 32 : c);
            set.insert(w);
        }
    }
}

const Words& wordsFor(const Lang* l) {
    static std::mutex m;
    static std::unordered_map<const Lang*, Words> cache;
    std::lock_guard<std::mutex> lock(m);
    auto it = cache.find(l);
    if (it != cache.end()) return it->second;
    Words& w = cache[l];
    const bool lower = (l->flags & F_CASELESS) != 0;
    addWords(w.keywords, l->keywords, lower);
    addWords(w.types, l->types, lower);
    addWords(w.literals, l->literals, lower);
    addWords(w.macros, l->macros, lower);
    return w;
}

std::string lowerAscii(std::string s) {
    for (char& c : s) c = (char)((c >= 'A' && c <= 'Z') ? c + 32 : c);
    return s;
}

}  // namespace

const char* langName(const Lang* l) { return l ? l->name : "Plain text"; }

const Lang* langForFile(const std::string& fileNameUtf8, const std::string& firstLine) {
    std::string name = lowerAscii(fileNameUtf8);
    size_t slash = name.find_last_of("\\/");
    if (slash != std::string::npos) name = name.substr(slash + 1);
    size_t dot = name.find_last_of('.');
    std::string ext = dot == std::string::npos ? std::string() : name.substr(dot + 1);
    for (const Lang& l : kLangs) {
        // Whole file names first ("=makefile"), then extensions.
        const char* p = l.exts;
        while (*p) {
            while (*p == ' ') p++;
            const char* s = p;
            while (*p && *p != ' ') p++;
            std::string w(s, (size_t)(p - s));
            if (w.empty()) continue;
            if (w[0] == '=') {
                if (name == w.substr(1)) return &l;
            } else if (!ext.empty() && ext == w) {
                return &l;
            }
        }
    }
    // A "#!" line names the interpreter.
    if (firstLine.size() > 2 && firstLine[0] == '#' && firstLine[1] == '!') {
        std::string f = lowerAscii(firstLine);
        struct { const char* word; const char* lang; } shebang[] = {
            {"python", "Python"}, {"bash", "Shell"}, {"/sh", "Shell"}, {"zsh", "Shell"}, {"node", "JavaScript"},
            {"ruby", "Ruby"},     {"perl", "Perl"},  {"php", "PHP"},   {"lua", "Lua"},   {"pwsh", "PowerShell"},
        };
        for (auto& s : shebang)
            if (f.find(s.word) != std::string::npos)
                for (const Lang& l : kLangs)
                    if (strcmp(l.name, s.lang) == 0) return &l;
    }
    if (firstLine.rfind("<?xml", 0) == 0)
        for (const Lang& l : kLangs)
            if (strcmp(l.name, "XML") == 0) return &l;
    return nullptr;
}

// ---------------------------------------------------------------- tokenizer

namespace {

// What carries over to the next line (low byte), plus details above it.
enum Mode : u32 {
    M_NORMAL = 0,
    M_BLOCK = 1,      // block comment; bits 8..15: nesting depth
    M_TRIPLE_DQ = 2,  // """
    M_TRIPLE_SQ = 3,  // '''
    M_BACKTICK = 4,   // `
    M_RAW_CPP = 5,    // R"( ... )"
    M_VERBATIM = 6,   // @"..."
    M_LUA_STR = 7,    // [[ ]]; bits 8..15: number of "="
    M_LUA_CMT = 8,    // --[[ ]]
    M_TAG = 9,        // inside <tag ...>
    M_MCOMMENT = 10,  // <!-- -->
    M_FENCE = 11,     // ``` code block in Markdown
    M_DQ = 12,        // "..." continuing over lines
    M_SQ = 13,        // '...' continuing over lines
};

inline bool isIdStart(u8 c) { return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || c == '_' || c >= 0x80; }
inline bool isDigit(u8 c) { return c >= '0' && c <= '9'; }
inline bool isIdChar(u8 c) { return isIdStart(c) || isDigit(c); }

struct Lexer {
    const Lang* L;
    const Words& W;
    const u8* p;
    size_t n;
    u8* k;  // may be null
    size_t i = 0;

    void mark(size_t a, size_t b, u8 kind) {
        if (!k) return;
        if (b > n) b = n;
        for (size_t j = a; j < b; j++) k[j] = kind;
    }
    bool at(const char* s) const {
        if (!s) return false;
        size_t m = strlen(s);
        return m && i + m <= n && memcmp(p + i, s, m) == 0;
    }
    bool atCi(const char* s) const {
        size_t m = strlen(s);
        if (i + m > n) return false;
        for (size_t j = 0; j < m; j++) {
            u8 c = p[i + j];
            if (c >= 'A' && c <= 'Z') c = (u8)(c + 32);
            if (c != (u8)s[j]) return false;
        }
        return true;
    }
    size_t skipSpace(size_t j) const {
        while (j < n && (p[j] == ' ' || p[j] == '\t')) j++;
        return j;
    }

    // Finds `close` from i; marks up to and including it. Returns true if found.
    bool until(const char* close, u8 kind) {
        size_t m = strlen(close);
        size_t s = i;
        while (i + m <= n) {
            if (memcmp(p + i, close, m) == 0) {
                i += m;
                mark(s, i, kind);
                return true;
            }
            i++;
        }
        i = n;
        mark(s, n, kind);
        return false;
    }

    // A quoted string from i (at the opening quote). Returns true if closed.
    bool quoted(u8 q, u8 kind, bool escapes = true) {
        size_t s = i++;
        while (i < n) {
            if (escapes && p[i] == '\\' && i + 1 < n) {
                i += 2;
                continue;
            }
            if (p[i] == q) {
                i++;
                mark(s, i, kind);
                return true;
            }
            i++;
        }
        mark(s, n, kind);
        return false;
    }
    // Continues a string from the start of the line (state carried over).
    bool continueQuoted(u8 q) {
        size_t s = i;
        while (i < n) {
            if (p[i] == '\\' && i + 1 < n) {
                i += 2;
                continue;
            }
            if (p[i] == q) {
                i++;
                mark(s, i, HL_STRING);
                return true;
            }
            i++;
        }
        mark(s, n, HL_STRING);
        return false;
    }

    void number() {
        size_t s = i;
        while (i < n) {
            u8 c = p[i];
            if (isIdChar(c) || c == '.' || c == '\'') {
                i++;
                continue;
            }
            if ((c == '+' || c == '-') && i > s && (p[i - 1] == 'e' || p[i - 1] == 'E' || p[i - 1] == 'p' || p[i - 1] == 'P') &&
                !(p[s] == '0' && s + 1 < n && (p[s + 1] == 'x' || p[s + 1] == 'X'))) {
                i++;
                continue;
            }
            break;
        }
        mark(s, i, HL_NUMBER);
    }

    // An identifier or keyword from i.
    void word() {
        size_t s = i;
        const u32 f = L->flags;
        while (i < n && (isIdChar(p[i]) || ((f & F_DASH_IDENT) && p[i] == '-' && i + 1 < n && isIdChar(p[i + 1])) ||
                         ((f & F_DOLLAR_IDENT) && p[i] == '$')))
            i++;
        std::string w((const char*)p + s, i - s);
        if (f & F_CASELESS) w = lowerAscii(w);
        u8 kind = HL_TEXT;
        if (W.keywords.count(w)) kind = HL_KEYWORD;
        else if (W.macros.count(w)) kind = HL_PREPROC;
        else if (W.types.count(w)) kind = HL_TYPE;
        else if (W.literals.count(w)) kind = HL_LITERAL;
        else {
            size_t j = skipSpace(i);
            if (j < n && p[j] == '(') kind = HL_FUNCTION;
            else if ((f & F_CSS) && j < n && p[j] == ':') kind = HL_ATTR;  // property name
        }
        mark(s, i, kind);
    }
};

}  // namespace

u32 highlightLine(const Lang* l, const u8* p, size_t n, u32 state, u8* kinds) {
    if (kinds) memset(kinds, HL_TEXT, n);
    if (!l) return 0;
    Lexer x{l, wordsFor(l), p, n, kinds};
    const u32 f = l->flags;
    u32 mode = state & 0xFF, arg = (state >> 8) & 0xFF;

    // ---- whole-line formats
    if (f & F_DIFF) {
        if (n && (p[0] == '+' || p[0] == '>') && !(n >= 3 && p[1] == '+' && p[2] == '+')) x.mark(0, n, HL_INSERTED);
        else if (n && (p[0] == '-' || p[0] == '<') && !(n >= 3 && p[1] == '-' && p[2] == '-')) x.mark(0, n, HL_DELETED);
        else if (n && p[0] == '@') x.mark(0, n, HL_PREPROC);
        else if (n >= 4 && (memcmp(p, "diff", 4) == 0 || memcmp(p, "+++", 3) == 0 || memcmp(p, "---", 3) == 0 ||
                            memcmp(p, "Inde", 4) == 0))
            x.mark(0, n, HL_KEYWORD);
        return 0;
    }
    if (f & F_MARKDOWN) {
        size_t j = x.skipSpace(0);
        if (j + 3 <= n && (memcmp(p + j, "```", 3) == 0 || memcmp(p + j, "~~~", 3) == 0)) {
            x.mark(0, n, HL_PREPROC);
            return mode == M_FENCE ? M_NORMAL : M_FENCE;
        }
        if (mode == M_FENCE) {
            x.mark(0, n, HL_STRING);
            return M_FENCE;
        }
        if (j < n && p[j] == '#') {
            x.mark(0, n, HL_TAG);
            return 0;
        }
        if (j < n && p[j] == '>') {
            x.mark(0, n, HL_COMMENT);
            return 0;
        }
        if (j < n && (p[j] == '-' || p[j] == '*' || p[j] == '+') && j + 1 < n && p[j + 1] == ' ') x.mark(j, j + 1, HL_KEYWORD);
        size_t d = j;
        while (d < n && isDigit(p[d])) d++;
        if (d > j && d + 1 < n && p[d] == '.' && p[d + 1] == ' ') x.mark(j, d + 1, HL_KEYWORD);
        for (size_t i = 0; i < n; i++) {
            if (p[i] == '`') {
                size_t e = i + 1;
                while (e < n && p[e] != '`') e++;
                x.mark(i, e + 1, HL_STRING);
                i = e;
            } else if (p[i] == '[') {  // [text](link)
                size_t e = i + 1;
                while (e < n && p[e] != ']') e++;
                if (e + 1 < n && p[e + 1] == '(') {
                    size_t u = e + 2;
                    while (u < n && p[u] != ')') u++;
                    x.mark(i, e + 1, HL_KEYWORD);
                    x.mark(e + 1, u + 1, HL_ATTR);
                    i = u;
                }
            } else if ((p[i] == '*' || p[i] == '_') && i + 1 < n && p[i + 1] == p[i]) {  // **bold**
                size_t e = i + 2;
                while (e + 1 < n && !(p[e] == p[i] && p[e + 1] == p[i])) e++;
                if (e + 1 < n) {
                    x.mark(i, e + 2, HL_TYPE);
                    i = e + 1;
                }
            }
        }
        return 0;
    }
    if (f & F_INI) {
        if (mode == M_TRIPLE_DQ || mode == M_TRIPLE_SQ) {
            if (!x.until(mode == M_TRIPLE_DQ ? "\"\"\"" : "'''", HL_STRING)) return mode;
            mode = M_NORMAL;
        }
        size_t j = x.skipSpace(x.i);
        if (x.i == 0 && j < n && (x.at(l->line1) || (p[j] == ';' || p[j] == '#'))) {
            x.mark(j, n, HL_COMMENT);
            return 0;
        }
        if (x.i == 0 && j < n && p[j] == '[') {
            x.mark(j, n, HL_TAG);
            return 0;
        }
        // key = value
        size_t eq = j;
        while (eq < n && p[eq] != '=' && p[eq] != ':' && p[eq] != '"') eq++;
        if (x.i == 0 && eq < n && p[eq] != '"') {
            x.mark(j, eq, HL_ATTR);
            x.i = eq + 1;
        }
        while (x.i < n) {
            u8 c = p[x.i];
            if ((f & F_TRIPLE) && (x.at("\"\"\"") || x.at("'''"))) {
                const char* q = x.at("\"\"\"") ? "\"\"\"" : "'''";
                size_t s = x.i;
                x.i += 3;
                if (!x.until(q, HL_STRING)) {
                    x.mark(s, n, HL_STRING);
                    return q[0] == '"' ? M_TRIPLE_DQ : M_TRIPLE_SQ;
                }
                x.mark(s, x.i, HL_STRING);
            } else if (c == '"' || (c == '\'' && (f & F_TRIPLE))) {
                x.quoted(c, HL_STRING, c == '"');
            } else if ((c == '#' || (c == ';' && !(f & F_TRIPLE))) && x.i > 0 &&
                       (p[x.i - 1] == ' ' || p[x.i - 1] == '\t')) {  // comment after a value
                x.mark(x.i, n, HL_COMMENT);
                break;
            } else if (isDigit(c) || (c == '-' && x.i + 1 < n && isDigit(p[x.i + 1]))) {
                x.i++;
                x.number();
            } else if (isIdStart(c)) {
                x.word();
            } else {
                x.i++;
            }
        }
        return 0;
    }
    if (f & F_YAML) {
        size_t j = x.skipSpace(0);
        if (j < n && p[j] == '#') {
            x.mark(j, n, HL_COMMENT);
            return 0;
        }
        if (n >= 3 && (memcmp(p, "---", 3) == 0 || memcmp(p, "...", 3) == 0)) {
            x.mark(0, n, HL_PREPROC);
            return 0;
        }
        if (j + 1 < n && p[j] == '-' && p[j + 1] == ' ') {
            x.mark(j, j + 1, HL_KEYWORD);
            j = x.skipSpace(j + 1);
        }
        // key:
        size_t c = j;
        if (c < n && (p[c] == '"' || p[c] == '\'')) {
            u8 q = p[c++];
            while (c < n && p[c] != q) c++;
            c++;
        } else {
            while (c < n && p[c] != ':' && p[c] != '#') c++;
        }
        x.i = j;
        if (c < n && p[c] == ':' && (c + 1 == n || p[c + 1] == ' ')) {
            x.mark(j, c, HL_ATTR);
            x.i = c + 1;
        }
        while (x.i < n) {
            u8 ch = p[x.i];
            if (ch == '#' && x.i > 0 && (p[x.i - 1] == ' ' || p[x.i - 1] == '\t')) {
                x.mark(x.i, n, HL_COMMENT);
                break;
            }
            if (ch == '"' || ch == '\'') x.quoted(ch, HL_STRING, ch == '"');
            else if (ch == '&' || ch == '*' || ch == '!') {  // anchors, aliases, tags
                size_t s = x.i++;
                while (x.i < n && p[x.i] != ' ') x.i++;
                x.mark(s, x.i, HL_PREPROC);
            } else if (isDigit(ch) || (ch == '-' && x.i + 1 < n && isDigit(p[x.i + 1]))) {
                x.i++;
                x.number();
            } else if (isIdStart(ch) || ch == '~') {
                size_t s = x.i;
                x.word();
                // Literals only when they are the whole value.
                std::string w = lowerAscii(std::string((const char*)p + s, x.i - s));
                if (!x.W.literals.count(w)) x.mark(s, x.i, HL_TEXT);
                if (ch == '~') x.i = s + 1;
            } else {
                x.i++;
            }
        }
        return 0;
    }
    if (f & F_BATCH) {
        size_t j = x.skipSpace(0);
        if (j + 1 < n && p[j] == '@') j++;
        if (x.atCi("rem") || (j + 3 <= n && (p[j] | 0x20) == 'r' && (p[j + 1] | 0x20) == 'e' && (p[j + 2] | 0x20) == 'm' &&
                              (j + 3 == n || p[j + 3] == ' '))) {
            x.mark(j, n, HL_COMMENT);
            return 0;
        }
        if (j + 1 < n && p[j] == ':' && p[j + 1] == ':') {
            x.mark(j, n, HL_COMMENT);
            return 0;
        }
        if (j < n && p[j] == ':') {
            x.mark(j, n, HL_TAG);  // label
            return 0;
        }
        while (x.i < n) {
            u8 c = p[x.i];
            if (c == '%') {
                size_t s = x.i++;
                if (x.i < n && p[x.i] == '%') x.i++;
                while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == '~')) x.i++;
                if (x.i < n && p[x.i] == '%') x.i++;
                x.mark(s, x.i, HL_LITERAL);
            } else if (c == '"') {
                x.quoted('"', HL_STRING, false);
            } else if (isDigit(c)) {
                x.number();
            } else if (isIdStart(c)) {
                x.word();
            } else {
                x.i++;
            }
        }
        return 0;
    }

    // ---- carried-over state
    switch (mode) {
    case M_BLOCK: {
        size_t s = x.i;
        int depth = arg ? arg : 1;
        while (x.i < n) {
            if (l->nested && x.at(l->blockOpen)) {
                depth++;
                x.i += strlen(l->blockOpen);
            } else if (x.at(l->blockClose)) {
                x.i += strlen(l->blockClose);
                if (--depth == 0) break;
            } else {
                x.i++;
            }
        }
        x.mark(s, x.i, HL_COMMENT);
        if (depth > 0) return M_BLOCK | ((u32)std::min(depth, 255) << 8);
        break;
    }
    case M_TRIPLE_DQ:
        if (!x.until("\"\"\"", HL_STRING)) return mode;
        break;
    case M_TRIPLE_SQ:
        if (!x.until("'''", HL_STRING)) return mode;
        break;
    case M_BACKTICK:
        if (!x.continueQuoted('`')) return mode;
        break;
    case M_RAW_CPP:
        if (!x.until(")\"", HL_STRING)) return mode;
        break;
    case M_VERBATIM: {
        size_t s = x.i;
        bool closed = false;
        while (x.i < n) {
            if (p[x.i] == '"') {
                if (x.i + 1 < n && p[x.i + 1] == '"') {
                    x.i += 2;
                    continue;
                }
                x.i++;
                closed = true;
                break;
            }
            x.i++;
        }
        x.mark(s, x.i, HL_STRING);
        if (!closed) return mode;
        break;
    }
    case M_LUA_STR:
    case M_LUA_CMT: {
        std::string close = "]" + std::string(arg, '=') + "]";
        if (!x.until(close.c_str(), mode == M_LUA_STR ? HL_STRING : HL_COMMENT)) return state;
        break;
    }
    case M_MCOMMENT:
        if (!x.until("-->", HL_COMMENT)) return mode;
        break;
    case M_DQ:
        if (!x.continueQuoted('"')) return mode;
        break;
    case M_SQ:
        if (!x.continueQuoted('\'')) return mode;
        break;
    case M_TAG:
        break;  // handled by the markup loop below
    }

    // ---- markup (XML / HTML)
    if (f & F_MARKUP) {
        bool inTag = mode == M_TAG;
        while (x.i < n) {
            u8 c = p[x.i];
            if (inTag) {
                if (c == '>' || (c == '/' && x.i + 1 < n && p[x.i + 1] == '>') || (c == '?' && x.i + 1 < n && p[x.i + 1] == '>')) {
                    size_t s = x.i;
                    x.i += c == '>' ? 1 : 2;
                    x.mark(s, x.i, HL_TAG);
                    inTag = false;
                } else if (c == '"' || c == '\'') {
                    x.quoted(c, HL_STRING, false);
                } else if (isIdStart(c) || c == ':' || c == '-') {
                    size_t s = x.i;
                    while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == ':' || p[x.i] == '-' || p[x.i] == '.')) x.i++;
                    x.mark(s, x.i, HL_ATTR);
                } else {
                    x.i++;
                }
                continue;
            }
            if (x.at("<!--")) {
                x.i += 4;
                x.mark(x.i - 4, x.i, HL_COMMENT);
                if (!x.until("-->", HL_COMMENT)) return M_MCOMMENT;
            } else if (x.at("<![CDATA[")) {
                size_t s = x.i;
                x.i += 9;
                x.until("]]>", HL_STRING);
                x.mark(s, s + 9, HL_PREPROC);
            } else if (c == '<' && x.i + 1 < n && (isIdStart(p[x.i + 1]) || p[x.i + 1] == '/' || p[x.i + 1] == '!' || p[x.i + 1] == '?')) {
                size_t s = x.i++;
                if (x.i < n && (p[x.i] == '/' || p[x.i] == '!' || p[x.i] == '?')) x.i++;
                while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == ':' || p[x.i] == '-' || p[x.i] == '.')) x.i++;
                x.mark(s, x.i, HL_TAG);
                inTag = true;
            } else if (c == '&') {
                size_t s = x.i++;
                while (x.i < n && x.i - s < 12 && (isIdChar(p[x.i]) || p[x.i] == '#')) x.i++;
                if (x.i < n && p[x.i] == ';') {
                    x.i++;
                    x.mark(s, x.i, HL_LITERAL);
                }
            } else {
                x.i++;
            }
        }
        return inTag ? M_TAG : M_NORMAL;
    }

    // ---- code
    // A preprocessor line: "#" first on the line (C family, shaders, C#, VB).
    if ((f & F_PREPROC_HASH) && mode == M_NORMAL && x.i == 0) {
        size_t j = x.skipSpace(0);
        if (j < n && p[j] == '#' && !(j + 1 < n && p[j + 1] == '!')) {
            size_t s = j++;
            j = x.skipSpace(j);
            while (j < n && isIdChar(p[j])) j++;
            x.mark(s, j, HL_PREPROC);
            // #include <file> / "file"
            size_t a = x.skipSpace(j);
            if (a < n && p[a] == '<') {
                size_t e = a;
                while (e < n && p[e] != '>') e++;
                x.mark(a, e + 1, HL_STRING);
                j = e + 1;
            }
            x.i = j;
        }
    }
    if (f & F_ASM) {
        // label:  and .directive
        size_t j = x.skipSpace(0);
        size_t e = j;
        while (e < n && (isIdChar(p[e]) || p[e] == '.' || p[e] == '@' || p[e] == '$')) e++;
        if (e > j && e < n && p[e] == ':') x.mark(j, e + 1, HL_TAG);
        else if (j < n && p[j] == '.') x.mark(j, e, HL_PREPROC);
    }

    while (x.i < n) {
        const u8 c = p[x.i];
        // comments (block first: Julia's "#=" also starts with "#")
        if (l->blockOpen && x.at(l->blockOpen)) {
            size_t s = x.i;
            x.i += strlen(l->blockOpen);
            int depth = 1;
            while (x.i < n) {
                if (l->nested && x.at(l->blockOpen)) {
                    depth++;
                    x.i += strlen(l->blockOpen);
                } else if (x.at(l->blockClose)) {
                    x.i += strlen(l->blockClose);
                    if (--depth == 0) break;
                } else {
                    x.i++;
                }
            }
            x.mark(s, x.i, HL_COMMENT);
            if (depth > 0) return M_BLOCK | ((u32)std::min(depth, 255) << 8);
            continue;
        }
        if (x.at(l->line1) || x.at(l->line2)) {
            if ((f & F_LUA) && x.at("--[")) {
                size_t j = x.i + 3, eq = 0;
                while (j < n && p[j] == '=') j++, eq++;
                if (j < n && p[j] == '[') {
                    size_t s = x.i;
                    x.i = j + 1;
                    std::string close = "]" + std::string(eq, '=') + "]";
                    if (!x.until(close.c_str(), HL_COMMENT)) {
                        x.mark(s, n, HL_COMMENT);
                        return M_LUA_CMT | ((u32)std::min<size_t>(eq, 255) << 8);
                    }
                    x.mark(s, x.i, HL_COMMENT);
                    continue;
                }
            }
            // "#" in PowerShell starts "<#" block comments too; "<#" is checked below first.
            x.mark(x.i, n, HL_COMMENT);
            return M_NORMAL;
        }
        // strings
        if ((f & F_TRIPLE) && (x.at("\"\"\"") || x.at("'''"))) {
            const char* q = x.at("\"\"\"") ? "\"\"\"" : "'''";
            size_t s = x.i;
            x.i += 3;
            if (!x.until(q, HL_STRING)) {
                x.mark(s, n, HL_STRING);
                return q[0] == '"' ? M_TRIPLE_DQ : M_TRIPLE_SQ;
            }
            x.mark(s, x.i, HL_STRING);
            continue;
        }
        if ((f & F_CPP_RAW) && c == 'R' && x.i + 1 < n && p[x.i + 1] == '"' &&
            (x.i == 0 || !isIdChar(p[x.i - 1]) || p[x.i - 1] == 'u' || p[x.i - 1] == 'U' || p[x.i - 1] == 'L' || p[x.i - 1] == '8')) {
            size_t s = x.i;
            x.i += 2;
            if (!x.until(")\"", HL_STRING)) {
                x.mark(s, n, HL_STRING);
                return M_RAW_CPP;
            }
            x.mark(s, x.i, HL_STRING);
            continue;
        }
        if ((f & F_CS_VERBATIM) && (x.at("@\"") || x.at("$@\"") || x.at("@$\""))) {
            size_t s = x.i;
            while (p[x.i] != '"') x.i++;
            x.i++;
            bool closed = false;
            while (x.i < n) {
                if (p[x.i] == '"') {
                    if (x.i + 1 < n && p[x.i + 1] == '"') {
                        x.i += 2;
                        continue;
                    }
                    x.i++;
                    closed = true;
                    break;
                }
                x.i++;
            }
            x.mark(s, x.i, HL_STRING);
            if (!closed) return M_VERBATIM;
            continue;
        }
        if ((f & F_LUA) && c == '[' && x.i + 1 < n && (p[x.i + 1] == '[' || p[x.i + 1] == '=')) {
            size_t j = x.i + 1, eq = 0;
            while (j < n && p[j] == '=') j++, eq++;
            if (j < n && p[j] == '[') {
                size_t s = x.i;
                x.i = j + 1;
                std::string close = "]" + std::string(eq, '=') + "]";
                if (!x.until(close.c_str(), HL_STRING)) {
                    x.mark(s, n, HL_STRING);
                    return M_LUA_STR | ((u32)std::min<size_t>(eq, 255) << 8);
                }
                x.mark(s, x.i, HL_STRING);
                continue;
            }
        }
        if ((f & F_BACKTICK) && c == '`') {
            if (!x.quoted('`', HL_STRING)) return M_BACKTICK;
            continue;
        }
        if (c == '\'' && (f & F_LIFETIMES)) {
            // 'x' or '\n' is a character; 'a without a closing quote is a lifetime.
            bool ch = (x.i + 2 < n && p[x.i + 2] == '\'') || (x.i + 1 < n && p[x.i + 1] == '\\');
            if (!ch) {
                size_t s = x.i++;
                while (x.i < n && isIdChar(p[x.i])) x.i++;
                x.mark(s, x.i, HL_TYPE);
                continue;
            }
        }
        if (l->quotes && c && strchr(l->quotes, c)) {
            const bool closed = x.quoted(c, HL_STRING, !(l->flags & F_CASELESS) || c == '"');
            if (!closed && (f & F_MULTILINE_STR)) return c == '"' ? M_DQ : M_SQ;
            continue;
        }
        // JSON keys: a string followed by ":"
        // (handled after the fact: the string was just marked)
        if (isDigit(c) || (c == '.' && x.i + 1 < n && isDigit(p[x.i + 1]))) {
            if (x.i > 0 && isIdChar(p[x.i - 1])) {
                x.i++;
                continue;
            }
            x.number();
            continue;
        }
        if ((f & F_DOLLAR_VARS) && c == '$') {
            size_t s = x.i++;
            if (x.i < n && (p[x.i] == '{' || p[x.i] == '(')) {
                u8 close = p[x.i] == '{' ? '}' : ')';
                while (x.i < n && p[x.i] != close) x.i++;
                if (x.i < n) x.i++;
            } else {
                while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == ':')) x.i++;
            }
            if (x.i == s + 1 && x.i < n && strchr("?!#@*0123456789_", p[x.i])) x.i++;
            std::string w = lowerAscii(std::string((const char*)p + s, x.i - s));
            x.mark(s, x.i, x.W.literals.count(w) ? HL_LITERAL : HL_ATTR);
            continue;
        }
        if ((f & F_AT_ANNOT) && c == '@' && x.i + 1 < n && isIdStart(p[x.i + 1])) {
            size_t s = x.i++;
            while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == '.')) x.i++;
            x.mark(s, x.i, HL_PREPROC);
            continue;
        }
        if ((f & F_CSS) && (c == '@' || c == '!')) {
            size_t s = x.i++;
            while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == '-')) x.i++;
            x.mark(s, x.i, HL_KEYWORD);
            continue;
        }
        if ((f & F_CSS) && c == '#' && x.i + 1 < n && isIdChar(p[x.i + 1])) {  // #id or #colour
            size_t s = x.i++;
            while (x.i < n && (isIdChar(p[x.i]) || p[x.i] == '-')) x.i++;
            x.mark(s, x.i, HL_NUMBER);
            continue;
        }
        if (isIdStart(c) || ((f & F_DOLLAR_IDENT) && c == '$') || ((f & F_DASH_IDENT) && c == '-' && x.i + 1 < n && isIdStart(p[x.i + 1]) &&
                                                                     (x.i == 0 || !isIdChar(p[x.i - 1])))) {
            x.word();
            continue;
        }
        x.i++;
    }
    // JSON keys: strings directly followed by ":".
    if ((f & F_JSON) && kinds) {
        for (size_t i = 0; i < n; i++) {
            if (kinds[i] != HL_STRING || (i > 0 && kinds[i - 1] == HL_STRING)) continue;
            size_t e = i;
            while (e < n && kinds[e] == HL_STRING) e++;
            size_t j = e;
            while (j < n && (p[j] == ' ' || p[j] == '\t')) j++;
            if (j < n && p[j] == ':') x.mark(i, e, HL_ATTR);
            i = e;
        }
    }
    return M_NORMAL;
}

}  // namespace jnp

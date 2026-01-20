using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using Mono.Cecil;
using Mono.Cecil.Cil;

namespace DllAnalyzer
{
    public class DllAnalyzer
    {
        private string? _currentAssemblyName;

        public DllInfo AnalyzeDll(string dllPath)
        {
            var readerParams = new ReaderParameters
            {
                ReadWrite = false,
                InMemory = true,
                ReadSymbols = false
            };

            using (var assembly = AssemblyDefinition.ReadAssembly(dllPath, readerParams))
            {
                _currentAssemblyName = assembly.Name.Name;

                var dllInfo = new DllInfo
                {
                    AssemblyName = assembly.Name.Name,
                    FullName = assembly.FullName,
                    Version = assembly.Name.Version?.ToString(),
                    Culture = assembly.Name.Culture,
                    PublicKeyToken = GetPublicKeyToken(assembly.Name.PublicKeyToken),
                    TargetRuntime = assembly.MainModule.Runtime.ToString(),
                    TargetFramework = GetTargetFramework(assembly),
                    Architecture = assembly.MainModule.Architecture.ToString(),
                    Kind = assembly.MainModule.Kind.ToString(),
                    Location = dllPath,
                    Types = new List<TypeInfo>(),
                    TypeForms = new List<TypeInfo>(),
                    TypeClasses = new List<TypeInfo>(),
                    TypeStaticClasses = new List<TypeInfo>()
                };

                foreach (var type in assembly.MainModule.Types)
                {
                    if (type.Name.Contains("<") || type.Name.StartsWith("__"))
                        continue;

                    var typeCategory = GetTypeCategory(type);
                    var isForm = typeCategory == "Form";

                    var typeInfo = new TypeInfo
                    {
                        FullName = type.FullName,
                        Name = type.Name,
                        Namespace = type.Namespace,
                        TypeCategory = typeCategory,
                        IsPublic = type.IsPublic,
                        IsAbstract = type.IsAbstract,
                        IsSealed = type.IsSealed,
                        IsStatic = type.IsAbstract && type.IsSealed,
                        IsInterface = type.IsInterface,
                        IsEnum = type.IsEnum,
                        IsValueType = type.IsValueType,
                        BaseType = type.BaseType?.FullName,
                        Interfaces = type.Interfaces.Select(i => i.InterfaceType.FullName).ToList(),
                        Fields = GetFields(type, isForm),
                        Properties = GetProperties(type),
                        Methods = GetMethods(type),
                        Constructors = GetConstructors(type),
                        Events = GetEvents(type),
                        NestedTypes = type.NestedTypes.Select(nt => nt.FullName).ToList()
                    };

                    if (isForm)
                    {
                        typeInfo.FormText = GetFormText(type);
                    }

                    dllInfo.Types.Add(typeInfo);

                    if (typeCategory == "Form")
                        dllInfo.TypeForms.Add(typeInfo);
                    else if (typeCategory == "StaticClass")
                        dllInfo.TypeStaticClasses.Add(typeInfo);
                    else if (typeCategory == "Class")
                        dllInfo.TypeClasses.Add(typeInfo);
                }

                return dllInfo;
            }
        }

        private string? GetPublicKeyToken(byte[]? token)
        {
            if (token == null || token.Length == 0)
                return null;

            return BitConverter.ToString(token).Replace("-", "").ToLower();
        }

        private string? GetTargetFramework(AssemblyDefinition assembly)
        {
            var targetFrameworkAttr = assembly.CustomAttributes
                .FirstOrDefault(a => a.AttributeType.FullName == "System.Runtime.Versioning.TargetFrameworkAttribute");

            if (targetFrameworkAttr != null && targetFrameworkAttr.ConstructorArguments.Count > 0)
            {
                return targetFrameworkAttr.ConstructorArguments[0].Value?.ToString();
            }

            return assembly.MainModule.Runtime.ToString();
        }

        private string GetTypeCategory(TypeDefinition type)
        {
            if (type.IsInterface)
                return "Interface";
            if (type.IsEnum)
                return "Enum";
            if (type.IsValueType && !type.IsEnum)
                return "Struct";

            if (type.IsClass)
            {
                var baseType = type.BaseType;
                while (baseType != null)
                {
                    if (baseType.FullName == "System.Windows.Forms.Form")
                        return "Form";

                    try
                    {
                        var resolved = baseType.Resolve();
                        baseType = resolved?.BaseType;
                    }
                    catch
                    {
                        break;
                    }
                }

                if (type.IsAbstract && type.IsSealed)
                    return "StaticClass";

                return "Class";
            }

            return "Type";
        }

        private string? GetFormText(TypeDefinition type)
        {
            try
            {
                var initMethod = type.Methods.FirstOrDefault(m => m.Name == "InitializeComponent");
                if (initMethod != null && initMethod.HasBody)
                {
                    for (int i = 0; i < initMethod.Body.Instructions.Count; i++)
                    {
                        var instruction = initMethod.Body.Instructions[i];

                        if (instruction.OpCode.Code == Code.Callvirt)
                        {
                            var method = instruction.Operand as MethodReference;
                            if (method != null && method.Name == "set_Text")
                            {
                                var loadThisInstr = instruction.Previous?.Previous;
                                if (loadThisInstr != null && loadThisInstr.OpCode.Code == Code.Ldarg_0)
                                {
                                    var valueInstr = instruction.Previous;
                                    if (valueInstr != null)
                                    {
                                        string? text = ExtractValue(valueInstr);
                                        if (!string.IsNullOrEmpty(text))
                                            return text;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private List<FieldInfo> GetFields(TypeDefinition type, bool isForm = false)
        {
            var fields = new List<FieldInfo>();

            foreach (var field in type.Fields)
            {
                if (field.Name.Contains("k__BackingField"))
                    continue;

                var fieldInfo = new FieldInfo
                {
                    Name = field.Name,
                    FullName = $"{field.FieldType.FullName} {field.Name}",
                    IdFullName = $"{type.FullName}.{field.Name}",
                    Type = field.FieldType.FullName,
                    IsPublic = field.IsPublic,
                    IsPrivate = field.IsPrivate,
                    IsStatic = field.IsStatic,
                    IsReadOnly = field.IsInitOnly,
                    IsLiteral = field.IsLiteral,
                    ConstantValue = field.HasConstant ? field.Constant?.ToString() : null,
                    References = new List<MemberReference>()
                };

                if (isForm && IsControlType(field.FieldType))
                {
                    fieldInfo.ControlProperties = GetControlProperties(field);
                }

                AnalyzeFieldReferences(field, fieldInfo);

                fields.Add(fieldInfo);
            }

            return fields;
        }

        private void AnalyzeFieldReferences(FieldDefinition field, FieldInfo fieldInfo)
        {
            try
            {
                var declaringType = field.DeclaringType;

                foreach (var method in declaringType.Methods)
                {
                    if (!method.HasBody)
                        continue;

                    foreach (var instruction in method.Body.Instructions)
                    {
                        if (instruction.OpCode.Code == Code.Ldfld ||
                            instruction.OpCode.Code == Code.Ldflda ||
                            instruction.OpCode.Code == Code.Stfld)
                        {
                            var fieldRef = instruction.Operand as FieldReference;
                            if (fieldRef != null && fieldRef.Name == field.Name)
                            {
                                var refInfo = new MemberReference
                                {
                                    MemberType = "Field",
                                    MemberName = field.Name,
                                    MemberFullName = field.FullName,
                                    ReferencedIn = method.Name,
                                    ReferencedInFullName = GetMethodSignature(method),
                                    ReferenceType = instruction.OpCode.Code == Code.Stfld ? "Write" : "Read"
                                };

                                if (!fieldInfo.References!.Any(r =>
                                    r.ReferencedIn == refInfo.ReferencedIn &&
                                    r.ReferenceType == refInfo.ReferenceType))
                                {
                                    fieldInfo.References.Add(refInfo);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing field references for {field.Name}: {ex.Message}");
            }
        }

        private bool IsControlType(TypeReference typeRef)
        {
            if (typeRef == null) return false;

            var typeName = typeRef.FullName;

            if (typeName.StartsWith("System.Windows.Forms.") ||
                typeName.StartsWith("DevExpress.") ||
                typeName.StartsWith("Telerik.") ||
                typeName.StartsWith("Janus.Windows.") ||
                typeName.StartsWith("System.Web.UI.WebControls."))
            {
                return true;
            }

            try
            {
                var typeDef = typeRef.Resolve();
                if (typeDef != null && typeDef.BaseType != null)
                {
                    return IsControlType(typeDef.BaseType);
                }
            }
            catch { }

            return false;
        }

        private Dictionary<string, string> GetControlProperties(FieldDefinition field)
        {
            var properties = new Dictionary<string, string>();

            try
            {
                var declaringType = field.DeclaringType;
                var initMethod = declaringType.Methods.FirstOrDefault(m => m.Name == "InitializeComponent");

                if (initMethod != null && initMethod.HasBody)
                {
                    var instructions = initMethod.Body.Instructions;

                    for (int i = 0; i < instructions.Count; i++)
                    {
                        var instruction = instructions[i];

                        if (instruction.OpCode.Code == Code.Ldfld)
                        {
                            var targetField = instruction.Operand as FieldReference;
                            if (targetField != null && targetField.Name == field.Name)
                            {
                                for (int j = i + 1; j < instructions.Count && j < i + 20; j++)
                                {
                                    var nextInstr = instructions[j];

                                    if (nextInstr.OpCode.Code == Code.Callvirt)
                                    {
                                        var method = nextInstr.Operand as MethodReference;
                                        if (method != null && method.Name.StartsWith("set_"))
                                        {
                                            string propName = method.Name.Substring(4);

                                            var valueInstr = nextInstr.Previous;
                                            if (valueInstr != null)
                                            {
                                                string? value = ExtractValue(valueInstr);
                                                if (!string.IsNullOrEmpty(value))
                                                {
                                                    if (!properties.ContainsKey(propName))
                                                        properties[propName] = value;
                                                }
                                                else
                                                {
                                                    var valueInstr2 = valueInstr.Previous;
                                                    if (valueInstr2 != null)
                                                    {
                                                        value = ExtractValue(valueInstr2);
                                                        if (!string.IsNullOrEmpty(value) && !properties.ContainsKey(propName))
                                                            properties[propName] = value;
                                                    }
                                                }
                                            }
                                        }

                                        if (method == null || !method.Name.StartsWith("set_"))
                                            break;
                                    }
                                    else if (nextInstr.OpCode.Code == Code.Ldfld ||
                                             nextInstr.OpCode.Code == Code.Stfld)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting properties for {field.Name}: {ex.Message}");
            }

            return properties;
        }

        private string? ExtractValue(Instruction instruction)
        {
            if (instruction == null) return null;

            try
            {
                switch (instruction.OpCode.Code)
                {
                    case Code.Ldstr:
                        return instruction.Operand?.ToString();

                    case Code.Ldc_I4_S:
                        return instruction.Operand?.ToString();

                    case Code.Ldc_I4_0: return "0";
                    case Code.Ldc_I4_1: return "1";
                    case Code.Ldc_I4_2: return "2";
                    case Code.Ldc_I4_3: return "3";
                    case Code.Ldc_I4_4: return "4";
                    case Code.Ldc_I4_5: return "5";
                    case Code.Ldc_I4_6: return "6";
                    case Code.Ldc_I4_7: return "7";
                    case Code.Ldc_I4_8: return "8";
                    case Code.Ldc_I4_M1: return "-1";

                    case Code.Ldc_I4:
                        var intValue = instruction.Operand as int?;
                        if (intValue.HasValue)
                        {
                            return intValue.Value.ToString();
                        }
                        break;

                    case Code.Ldc_R4:
                    case Code.Ldc_R8:
                        return instruction.Operand?.ToString();

                    case Code.Ldc_I8:
                        return instruction.Operand?.ToString();

                    case Code.Ldsfld:
                        var fieldRef = instruction.Operand as FieldReference;
                        if (fieldRef != null)
                        {
                            return $"{fieldRef.DeclaringType.Name}.{fieldRef.Name}";
                        }
                        break;
                }
            }
            catch { }

            return null;
        }

        private List<PropertyInfo> GetProperties(TypeDefinition type)
        {
            var properties = new List<PropertyInfo>();

            foreach (var prop in type.Properties)
            {
                var getter = prop.GetMethod;
                var setter = prop.SetMethod;

                var propInfo = new PropertyInfo
                {
                    Name = prop.Name,
                    FullName = $"{prop.PropertyType.FullName} {prop.Name}",
                    IdFullName = $"{type.FullName}.{prop.Name}",
                    Type = prop.PropertyType.FullName,
                    CanRead = getter != null,
                    CanWrite = setter != null,
                    IsStatic = (getter ?? setter)?.IsStatic ?? false,
                    GetterVisibility = getter != null ? GetVisibility(getter) : null,
                    SetterVisibility = setter != null ? GetVisibility(setter) : null,
                    References = new List<MemberReference>()
                };

                AnalyzePropertyReferences(prop, propInfo, type);

                properties.Add(propInfo);
            }

            return properties;
        }

        private void AnalyzePropertyReferences(PropertyDefinition prop, PropertyInfo propInfo, TypeDefinition type)
        {
            try
            {
                foreach (var method in type.Methods)
                {
                    if (!method.HasBody)
                        continue;

                    foreach (var instruction in method.Body.Instructions)
                    {
                        if (instruction.OpCode.Code == Code.Call ||
                            instruction.OpCode.Code == Code.Callvirt)
                        {
                            var methodRef = instruction.Operand as MethodReference;
                            if (methodRef != null)
                            {
                                bool isGetter = prop.GetMethod != null && methodRef.Name == prop.GetMethod.Name;
                                bool isSetter = prop.SetMethod != null && methodRef.Name == prop.SetMethod.Name;

                                if (isGetter || isSetter)
                                {
                                    var refInfo = new MemberReference
                                    {
                                        MemberType = "Property",
                                        MemberName = prop.Name,
                                        MemberFullName = propInfo.FullName,
                                        MemberIdFullName = $"{type.FullName}.{prop.Name}",
                                        ReferencedIn = method.Name,
                                        ReferencedInFullName = GetMethodSignature(method),
                                        ReferenceType = isSetter ? "Write" : "Read"
                                    };

                                    if (!propInfo.References!.Any(r =>
                                        r.ReferencedIn == refInfo.ReferencedIn &&
                                        r.ReferenceType == refInfo.ReferenceType))
                                    {
                                        propInfo.References.Add(refInfo);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing property references for {prop.Name}: {ex.Message}");
            }
        }

        private string? GetVisibility(MethodDefinition method)
        {
            if (method.IsPublic) return "Public";
            if (method.IsPrivate) return "Private";
            if (method.IsFamily) return "Protected";
            if (method.IsAssembly) return "Internal";
            if (method.IsFamilyOrAssembly) return "ProtectedInternal";
            return "Unknown";
        }

        private List<MethodInfo> GetMethods(TypeDefinition type)
        {
            var methods = new List<MethodInfo>();

            foreach (var method in type.Methods)
            {
                if (method.IsSpecialName || method.IsConstructor)
                    continue;

                var parameters = method.Parameters;
                var paramStrings = parameters.Select(p => $"{p.ParameterType.FullName} {p.Name}").ToArray();
                var paramTypesOnly = parameters.Select(p => p.ParameterType.FullName).ToArray();
                var paramNamesOnly = parameters.Select(p => p.Name).ToArray();

                string paramSignature = string.Join(",", paramTypesOnly);

                var methodInfo = new MethodInfo
                {
                    Name = method.Name,
                    FullNameWithArgs = $"{method.ReturnType.FullName} {method.Name}({string.Join(", ", paramStrings)})",
                    FullNameWithoutArgs = $"{method.ReturnType.FullName} {method.Name}()",
                    IdFullName = $"{type.FullName}.{method.Name}({paramSignature})",
                    ReturnType = method.ReturnType.FullName,
                    Parameters = paramStrings.ToList(),
                    ParameterTypes = paramTypesOnly.ToList(),
                    ParameterNames = paramNamesOnly.ToList(),
                    Visibility = GetVisibility(method),
                    IsStatic = method.IsStatic,
                    IsVirtual = method.IsVirtual,
                    IsAbstract = method.IsAbstract,
                    IsFinal = method.IsFinal,
                    IsAsync = IsAsyncMethod(method),
                    References = new List<MemberReference>(),
                    HasBody = method.HasBody
                };

                AnalyzeMethodReferences(method, methodInfo, type);

                methods.Add(methodInfo);
            }

            return methods;
        }

        private void AnalyzeMethodReferences(MethodDefinition method, MethodInfo methodInfo, TypeDefinition type)
        {
            if (!method.HasBody)
                return;

            try
            {
                foreach (var instruction in method.Body.Instructions)
                {
                    if (instruction.OpCode.Code == Code.Call ||
                        instruction.OpCode.Code == Code.Callvirt ||
                        instruction.OpCode.Code == Code.Calli)
                    {
                        var methodRef = instruction.Operand as MethodReference;
                        if (methodRef != null)
                        {
                            var refInfo = new MemberReference
                            {
                                MemberType = "Method",
                                MemberName = methodRef.Name,
                                MemberFullName = GetMethodReferenceSignature(methodRef),
                                MemberIdFullName = GetMethodReferenceIdFullName(methodRef),
                                ReferencedIn = method.Name,
                                ReferencedInFullName = GetMethodSignature(method),
                                ReferenceType = "Call"
                            };

                            if (!methodRef.Name.StartsWith("get_") &&
                                !methodRef.Name.StartsWith("set_") &&
                                !methodRef.Name.StartsWith("add_") &&
                                !methodRef.Name.StartsWith("remove_"))
                            {
                                if (!methodInfo.References!.Any(r =>
                                    r.MemberIdFullName == refInfo.MemberIdFullName))
                                {
                                    methodInfo.References.Add(refInfo);
                                }
                            }
                        }
                    }

                    if (instruction.OpCode.Code == Code.Ldfld ||
                        instruction.OpCode.Code == Code.Ldflda ||
                        instruction.OpCode.Code == Code.Stfld ||
                        instruction.OpCode.Code == Code.Ldsfld ||
                        instruction.OpCode.Code == Code.Stsfld)
                    {
                        var fieldRef = instruction.Operand as FieldReference;
                        if (fieldRef != null)
                        {
                            var refInfo = new MemberReference
                            {
                                MemberType = "Field",
                                MemberName = fieldRef.Name,
                                MemberFullName = $"{fieldRef.FieldType.FullName} {fieldRef.Name}",
                                MemberIdFullName = $"{fieldRef.DeclaringType.FullName}.{fieldRef.Name}",
                                ReferencedIn = method.Name,
                                ReferencedInFullName = GetMethodSignature(method),
                                ReferenceType = (instruction.OpCode.Code == Code.Stfld || instruction.OpCode.Code == Code.Stsfld) ? "Write" : "Read"
                            };

                            if (!methodInfo.References!.Any(r =>
                                r.MemberIdFullName == refInfo.MemberIdFullName &&
                                r.ReferenceType == refInfo.ReferenceType))
                            {
                                methodInfo.References.Add(refInfo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing method references for {method.Name}: {ex.Message}");
            }
        }

        private string GetMethodSignature(MethodDefinition method)
        {
            var paramTypes = method.Parameters.Select(p => p.ParameterType.FullName);
            return $"{method.DeclaringType.FullName}.{method.Name}({string.Join(",", paramTypes)})";
        }

        private string GetMethodReferenceSignature(MethodReference methodRef)
        {
            var paramTypes = methodRef.Parameters.Select(p => p.ParameterType.FullName);
            return $"{methodRef.DeclaringType.FullName}.{methodRef.Name}({string.Join(",", paramTypes)})";
        }

        private string GetMethodReferenceIdFullName(MethodReference methodRef)
        {
            var paramTypes = methodRef.Parameters.Select(p => p.ParameterType.FullName);
            return $"{methodRef.DeclaringType.FullName}.{methodRef.Name}({string.Join(",", paramTypes)})";
        }

        private bool IsAsyncMethod(MethodDefinition method)
        {
            return method.CustomAttributes.Any(a =>
                a.AttributeType.FullName == "System.Runtime.CompilerServices.AsyncStateMachineAttribute");
        }

        private List<ConstructorInfo> GetConstructors(TypeDefinition type)
        {
            var constructors = new List<ConstructorInfo>();

            foreach (var ctor in type.Methods.Where(m => m.IsConstructor))
            {
                var parameters = ctor.Parameters;
                var paramStrings = parameters.Select(p => $"{p.ParameterType.FullName} {p.Name}").ToArray();
                var paramTypesOnly = parameters.Select(p => p.ParameterType.FullName).ToArray();

                string paramSignature = string.Join(",", paramTypesOnly);

                var ctorInfo = new ConstructorInfo
                {
                    FullNameWithArgs = $"{type.Name}({string.Join(", ", paramStrings)})",
                    FullNameWithoutArgs = $"{type.Name}()",
                    IdFullName = $"{type.FullName}..ctor({paramSignature})",
                    Parameters = paramStrings.ToList(),
                    Visibility = GetVisibility(ctor),
                    IsStatic = ctor.IsStatic,
                    References = new List<MemberReference>()
                };

                AnalyzeConstructorReferences(ctor, ctorInfo, type);

                constructors.Add(ctorInfo);
            }

            return constructors;
        }

        private void AnalyzeConstructorReferences(MethodDefinition method, ConstructorInfo ctorInfo, TypeDefinition type)
        {
            if (!method.HasBody)
                return;

            try
            {
                foreach (var instruction in method.Body.Instructions)
                {
                    if (instruction.OpCode.Code == Code.Call ||
                        instruction.OpCode.Code == Code.Callvirt)
                    {
                        var methodRef = instruction.Operand as MethodReference;
                        if (methodRef != null)
                        {
                            var refInfo = new MemberReference
                            {
                                MemberType = "Method",
                                MemberName = methodRef.Name,
                                MemberFullName = GetMethodReferenceSignature(methodRef),
                                MemberIdFullName = GetMethodReferenceIdFullName(methodRef),
                                ReferencedIn = method.Name,
                                ReferencedInFullName = GetMethodSignature(method),
                                ReferenceType = "Call"
                            };

                            if (!methodRef.Name.StartsWith("get_") &&
                                !methodRef.Name.StartsWith("set_") &&
                                !methodRef.Name.StartsWith("add_") &&
                                !methodRef.Name.StartsWith("remove_"))
                            {
                                if (!ctorInfo.References!.Any(r =>
                                    r.MemberIdFullName == refInfo.MemberIdFullName))
                                {
                                    ctorInfo.References.Add(refInfo);
                                }
                            }
                        }
                    }

                    if (instruction.OpCode.Code == Code.Ldfld ||
                        instruction.OpCode.Code == Code.Ldflda ||
                        instruction.OpCode.Code == Code.Stfld ||
                        instruction.OpCode.Code == Code.Ldsfld ||
                        instruction.OpCode.Code == Code.Stsfld)
                    {
                        var fieldRef = instruction.Operand as FieldReference;
                        if (fieldRef != null)
                        {
                            var refInfo = new MemberReference
                            {
                                MemberType = "Field",
                                MemberName = fieldRef.Name,
                                MemberFullName = $"{fieldRef.FieldType.FullName} {fieldRef.Name}",
                                MemberIdFullName = $"{fieldRef.DeclaringType.FullName}.{fieldRef.Name}",
                                ReferencedIn = method.Name,
                                ReferencedInFullName = GetMethodSignature(method),
                                ReferenceType = (instruction.OpCode.Code == Code.Stfld || instruction.OpCode.Code == Code.Stsfld) ? "Write" : "Read"
                            };

                            if (!ctorInfo.References!.Any(r =>
                                r.MemberIdFullName == refInfo.MemberIdFullName &&
                                r.ReferenceType == refInfo.ReferenceType))
                            {
                                ctorInfo.References.Add(refInfo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error analyzing constructor references: {ex.Message}");
            }
        }

        private List<EventInfo> GetEvents(TypeDefinition type)
        {
            var events = new List<EventInfo>();

            foreach (var evt in type.Events)
            {
                events.Add(new EventInfo
                {
                    Name = evt.Name,
                    EventType = evt.EventType.FullName,
                    FullName = $"{evt.EventType.FullName} {evt.Name}",
                    IdFullName = $"{type.FullName}.{evt.Name}"
                });
            }

            return events;
        }

        public void SaveToJson(DllInfo dllInfo, string outputPath)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string json = JsonSerializer.Serialize(dllInfo, options);
            File.WriteAllText(outputPath, json, System.Text.Encoding.UTF8);
        }

        public void SaveSummaryToJson(DllInfo dllInfo, string outputPath)
        {
            var summary = new DllSummaryInfo
            {
                AssemblyName = dllInfo.AssemblyName,
                FullName = dllInfo.FullName,
                Version = dllInfo.Version,
                Culture = dllInfo.Culture,
                TargetRuntime = dllInfo.TargetRuntime,
                TargetFramework = dllInfo.TargetFramework,
                Architecture = dllInfo.Architecture,
                Kind = dllInfo.Kind,
                TypeForms = dllInfo.TypeForms?.Select(t => CreateTypeSummary(t)).ToList(),
                TypeClasses = dllInfo.TypeClasses?.Select(t => CreateTypeSummary(t)).ToList(),
                TypeStaticClasses = dllInfo.TypeStaticClasses?.Select(t => CreateTypeSummary(t)).ToList()
            };

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string json = JsonSerializer.Serialize(summary, options);
            File.WriteAllText(outputPath, json, System.Text.Encoding.UTF8);
        }

        private TypeSummaryInfo CreateTypeSummary(TypeInfo type)
        {
            var summary = new TypeSummaryInfo
            {
                FullName = type.FullName,
                Name = type.Name,
                Namespace = type.Namespace,
                TypeCategory = type.TypeCategory,
                FormText = type.FormText,
                Members = new List<string>()
            };

            if (type.Fields != null)
            {
                foreach (var field in type.Fields)
                {
                    if (!string.IsNullOrEmpty(field.IdFullName))
                        summary.Members!.Add(field.IdFullName);
                }
            }

            if (type.Properties != null)
            {
                foreach (var prop in type.Properties)
                {
                    if (!string.IsNullOrEmpty(prop.IdFullName))
                        summary.Members!.Add(prop.IdFullName);
                }
            }

            if (type.Methods != null)
            {
                foreach (var method in type.Methods)
                {
                    if (!string.IsNullOrEmpty(method.IdFullName))
                        summary.Members!.Add(method.IdFullName);

                    var idWithoutArgs = $"{type.FullName}.{method.Name}()";
                    if (!summary.Members!.Contains(idWithoutArgs))
                        summary.Members!.Add(idWithoutArgs);
                }
            }

            if (type.Constructors != null)
            {
                foreach (var ctor in type.Constructors)
                {
                    if (!string.IsNullOrEmpty(ctor.IdFullName))
                        summary.Members!.Add(ctor.IdFullName);

                    var idWithoutArgs = $"{type.FullName}..ctor()";
                    if (!summary.Members!.Contains(idWithoutArgs))
                        summary.Members!.Add(idWithoutArgs);
                }
            }

            if (type.Events != null)
            {
                foreach (var evt in type.Events)
                {
                    if (!string.IsNullOrEmpty(evt.IdFullName))
                        summary.Members!.Add(evt.IdFullName);
                }
            }

            return summary;
        }

        public void SaveExternalReferencesToJson(DllInfo dllInfo, string outputPath)
        {
            var externalRefs = new ExternalReferencesInfo
            {
                AssemblyName = dllInfo.AssemblyName,
                FullName = dllInfo.FullName,
                Version = dllInfo.Version,
                ExternalCalls = new List<ExternalCallInfo>(),
                ExternalCallsIdFullName = new List<string>()
            };

            if (dllInfo.Types != null)
            {
                foreach (var type in dllInfo.Types)
                {
                    AnalyzeTypeForExternalCalls(type, externalRefs.ExternalCalls);
                }
            }

            // Extract unique IdFullName without args and ()
            var uniqueIdFullNames = externalRefs.ExternalCalls
                .Where(ec => !string.IsNullOrEmpty(ec.CalledMemberFullName))
                .Select(ec => RemoveMethodArgs(ec.CalledMemberFullName!))
                .Distinct()
                .OrderBy(id => id)
                .ToList();

            externalRefs.ExternalCallsIdFullName = uniqueIdFullNames;

            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            string json = JsonSerializer.Serialize(externalRefs, options);
            File.WriteAllText(outputPath, json, System.Text.Encoding.UTF8);
        }

        private string RemoveMethodArgs(string fullName)
        {
            // Remove args from: "Namespace.Class.Method(args)" -> "Namespace.Class.Method"
            var parenIndex = fullName.IndexOf('(');
            if (parenIndex > 0)
            {
                return fullName.Substring(0, parenIndex);
            }
            return fullName;
        }

        private void AnalyzeTypeForExternalCalls(TypeInfo type, List<ExternalCallInfo> externalCalls)
        {
            if (type.Methods == null) return;

            foreach (var method in type.Methods)
            {
                if (!method.HasBody || method.References == null) continue;

                var triggerInfo = FindMethodTrigger(type, method.Name);

                foreach (var reference in method.References)
                {
                    if (!string.IsNullOrEmpty(reference.MemberIdFullName))
                    {
                        var typeName = ExtractTypeFromIdFullName(reference.MemberIdFullName);

                        if (!IsInternalType(typeName) && IsRelevantExternalType(typeName))
                        {
                            var externalCall = new ExternalCallInfo
                            {
                                CalledType = typeName,
                                CalledMember = reference.MemberName,
                                CalledMemberFullName = reference.MemberIdFullName,
                                CalledMemberType = reference.MemberType,
                                CalledFrom = new CalledFromInfo
                                {
                                    TypeFullName = type.FullName,
                                    TypeName = type.Name,
                                    TypeCategory = type.TypeCategory,
                                    FormText = type.FormText,
                                    MethodName = method.Name,
                                    MethodFullName = method.IdFullName,
                                    TriggerControl = triggerInfo?.ControlName,
                                    TriggerControlText = triggerInfo?.ControlText,
                                    TriggerEvent = triggerInfo?.EventName
                                }
                            };

                            if (!externalCalls.Any(ec =>
                                ec.CalledType == externalCall.CalledType &&
                                ec.CalledMember == externalCall.CalledMember &&
                                ec.CalledFrom?.MethodFullName == externalCall.CalledFrom?.MethodFullName))
                            {
                                externalCalls.Add(externalCall);
                            }
                        }
                    }
                }
            }
        }

        private string ExtractTypeFromIdFullName(string idFullName)
        {
            var lastDotBeforeMethod = idFullName.LastIndexOf('.');
            if (lastDotBeforeMethod > 0)
            {
                var typePart = idFullName.Substring(0, lastDotBeforeMethod);
                var secondLastDot = typePart.LastIndexOf('.');
                if (secondLastDot > 0)
                {
                    return typePart;
                }
            }

            return idFullName;
        }

        private TriggerInfo? FindMethodTrigger(TypeInfo type, string? methodName)
        {
            if (string.IsNullOrEmpty(methodName) || type.Fields == null)
                return null;

            var parts = methodName.Split('_');
            if (parts.Length >= 2)
            {
                var controlName = parts[0];
                var eventName = string.Join("_", parts.Skip(1));

                var controlField = type.Fields.FirstOrDefault(f =>
                    f.Name == controlName && f.ControlProperties != null);

                if (controlField != null)
                {
                    var controlText = controlField.ControlProperties?.ContainsKey("Text") == true
                        ? controlField.ControlProperties["Text"]
                        : null;

                    return new TriggerInfo
                    {
                        ControlName = controlName,
                        ControlText = controlText,
                        EventName = eventName
                    };
                }
            }

            return null;
        }

        private bool IsInternalType(string typeName)
        {
            if (string.IsNullOrEmpty(typeName) || string.IsNullOrEmpty(_currentAssemblyName))
                return false;

            return typeName.StartsWith(_currentAssemblyName + ".");
        }

        private bool IsRelevantExternalType(string typeName)
        {
            // Skip System namespaces
            if (typeName.StartsWith("System."))
            {
                if (typeName.StartsWith("System.Windows.Forms.")) return false;
                if (typeName.StartsWith("System.Drawing.")) return false;
                if (typeName.StartsWith("System.Collections.")) return false;
                if (typeName.StartsWith("System.Linq.")) return false;
                if (typeName.StartsWith("System.Text.")) return false;
                if (typeName.StartsWith("System.IO.")) return false;
                if (typeName.StartsWith("System.Threading.")) return false;
                if (typeName.StartsWith("System.Data.")) return false;
                if (typeName.StartsWith("System.ComponentModel.")) return false;
                if (typeName.StartsWith("System.Runtime.")) return false;
                if (typeName.StartsWith("System.Reflection.")) return false;
                if (typeName.StartsWith("System.Diagnostics.")) return false;
                if (typeName.StartsWith("System.Net.")) return false;
                if (typeName.StartsWith("System.Xml.")) return false;
                return false;
            }

            // Skip Microsoft namespaces
            if (typeName.StartsWith("Microsoft.")) return false;

            // Skip common third-party libraries (không cần thiết)
            if (typeName.StartsWith("Newtonsoft.")) return false;
            if (typeName.StartsWith("RestSharp.")) return false;
            if (typeName.StartsWith("AutoMapper.")) return false;
            if (typeName.StartsWith("Dapper.")) return false;
            if (typeName.StartsWith("Serilog.")) return false;
            if (typeName.StartsWith("NLog.")) return false;
            if (typeName.StartsWith("FluentValidation.")) return false;
            if (typeName.StartsWith("MediatR.")) return false;
            if (typeName.StartsWith("EntityFramework.")) return false;
            if (typeName.StartsWith("Npgsql.")) return false;
            if (typeName.StartsWith("Oracle.")) return false;
            if (typeName.StartsWith("MySql.")) return false;
            if (typeName.StartsWith("SqlKata.")) return false;
            if (typeName.StartsWith("Janus.")) return false;
            if (typeName.StartsWith("DevExpress.")) return false;
            if (typeName.StartsWith("Telerik.")) return false;
            if (typeName.StartsWith("Npgsql.")) return false;

            // Include any other custom/internal types
            return true;
        }
    }

    // Data Models
    public class DllInfo
    {
        public string? AssemblyName { get; set; }
        public string? FullName { get; set; }
        public string? Version { get; set; }
        public string? Culture { get; set; }
        public string? PublicKeyToken { get; set; }
        public string? TargetRuntime { get; set; }
        public string? TargetFramework { get; set; }
        public string? Architecture { get; set; }
        public string? Kind { get; set; }
        public string? Location { get; set; }
        public List<TypeInfo>? Types { get; set; }
        public List<TypeInfo>? TypeForms { get; set; }
        public List<TypeInfo>? TypeClasses { get; set; }
        public List<TypeInfo>? TypeStaticClasses { get; set; }
    }

    public class TypeInfo
    {
        public string? FullName { get; set; }
        public string? Name { get; set; }
        public string? Namespace { get; set; }
        public string? TypeCategory { get; set; }
        public string? FormText { get; set; }
        public bool IsPublic { get; set; }
        public bool IsAbstract { get; set; }
        public bool IsSealed { get; set; }
        public bool IsStatic { get; set; }
        public bool IsInterface { get; set; }
        public bool IsEnum { get; set; }
        public bool IsValueType { get; set; }
        public string? BaseType { get; set; }
        public List<string>? Interfaces { get; set; }
        public List<string>? NestedTypes { get; set; }
        public List<FieldInfo>? Fields { get; set; }
        public List<PropertyInfo>? Properties { get; set; }
        public List<MethodInfo>? Methods { get; set; }
        public List<ConstructorInfo>? Constructors { get; set; }
        public List<EventInfo>? Events { get; set; }
    }

    public class MemberReference
    {
        public string? MemberType { get; set; }
        public string? MemberName { get; set; }
        public string? MemberFullName { get; set; }
        public string? MemberIdFullName { get; set; }
        public string? ReferencedIn { get; set; }
        public string? ReferencedInFullName { get; set; }
        public string? ReferenceType { get; set; }
    }

    public class FieldInfo
    {
        public string? Name { get; set; }
        public string? FullName { get; set; }
        public string? IdFullName { get; set; }
        public string? Type { get; set; }
        public bool IsPublic { get; set; }
        public bool IsPrivate { get; set; }
        public bool IsStatic { get; set; }
        public bool IsReadOnly { get; set; }
        public bool IsLiteral { get; set; }
        public string? ConstantValue { get; set; }
        public Dictionary<string, string>? ControlProperties { get; set; }
        public List<MemberReference>? References { get; set; }
    }

    public class PropertyInfo
    {
        public string? Name { get; set; }
        public string? FullName { get; set; }
        public string? IdFullName { get; set; }
        public string? Type { get; set; }
        public bool CanRead { get; set; }
        public bool CanWrite { get; set; }
        public bool IsStatic { get; set; }
        public string? GetterVisibility { get; set; }
        public string? SetterVisibility { get; set; }
        public List<MemberReference>? References { get; set; }
    }

    public class MethodInfo
    {
        public string? Name { get; set; }
        public string? FullNameWithArgs { get; set; }
        public string? FullNameWithoutArgs { get; set; }
        public string? IdFullName { get; set; }
        public string? ReturnType { get; set; }
        public List<string>? Parameters { get; set; }
        public List<string>? ParameterTypes { get; set; }
        public List<string>? ParameterNames { get; set; }
        public string? Visibility { get; set; }
        public bool IsStatic { get; set; }
        public bool IsVirtual { get; set; }
        public bool IsAbstract { get; set; }
        public bool IsFinal { get; set; }
        public bool IsAsync { get; set; }
        public bool HasBody { get; set; }
        public List<MemberReference>? References { get; set; }
    }

    public class ConstructorInfo
    {
        public string? FullNameWithArgs { get; set; }
        public string? FullNameWithoutArgs { get; set; }
        public string? IdFullName { get; set; }
        public List<string>? Parameters { get; set; }
        public string? Visibility { get; set; }
        public bool IsStatic { get; set; }
        public List<MemberReference>? References { get; set; }
    }

    public class EventInfo
    {
        public string? Name { get; set; }
        public string? EventType { get; set; }
        public string? FullName { get; set; }
        public string? IdFullName { get; set; }
    }

    public class DllSummaryInfo
    {
        public string? AssemblyName { get; set; }
        public string? FullName { get; set; }
        public string? Version { get; set; }
        public string? Culture { get; set; }
        public string? TargetRuntime { get; set; }
        public string? TargetFramework { get; set; }
        public string? Architecture { get; set; }
        public string? Kind { get; set; }
        public List<TypeSummaryInfo>? TypeForms { get; set; }
        public List<TypeSummaryInfo>? TypeClasses { get; set; }
        public List<TypeSummaryInfo>? TypeStaticClasses { get; set; }
    }

    public class TypeSummaryInfo
    {
        public string? FullName { get; set; }
        public string? Name { get; set; }
        public string? Namespace { get; set; }
        public string? TypeCategory { get; set; }
        public string? FormText { get; set; }
        public List<string>? Members { get; set; }
    }

    public class ExternalReferencesInfo
    {
        public string? AssemblyName { get; set; }
        public string? FullName { get; set; }
        public string? Version { get; set; }
        public List<string>? ExternalCallsIdFullName { get; set; }
        public List<ExternalCallInfo>? ExternalCalls { get; set; }
    }

    public class ExternalCallInfo
    {
        public string? CalledType { get; set; }
        public string? CalledMember { get; set; }
        public string? CalledMemberFullName { get; set; }
        public string? CalledMemberType { get; set; }
        public CalledFromInfo? CalledFrom { get; set; }
    }

    public class CalledFromInfo
    {
        public string? TypeFullName { get; set; }
        public string? TypeName { get; set; }
        public string? TypeCategory { get; set; }
        public string? FormText { get; set; }
        public string? MethodName { get; set; }
        public string? MethodFullName { get; set; }
        public string? TriggerControl { get; set; }
        public string? TriggerControlText { get; set; }
        public string? TriggerEvent { get; set; }
    }

    public class TriggerInfo
    {
        public string? ControlName { get; set; }
        public string? ControlText { get; set; }
        public string? EventName { get; set; }
    }
}
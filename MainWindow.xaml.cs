using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Xml;
using System.Xml.Linq;

namespace StyleSnooper
{
    public sealed partial class MainWindow : INotifyPropertyChanged
    {
        private readonly Style _bracketStyle, _elementStyle, _quotesStyle, _textStyle, _attributeStyle;

        public MainWindow()
        {
            Styles = GetStyles(typeof(FrameworkElement).Assembly).ToList();

            InitializeComponent();

            // get syntax coloring styles
            _bracketStyle = (Style)Resources["BracketStyle"];
            _elementStyle = (Style)Resources["ElementStyle"];
            _quotesStyle = (Style)Resources["QuotesStyle"];
            _textStyle = (Style)Resources["TextStyle"];
            _attributeStyle = (Style)Resources["AttributeStyle"];

            // start out by looking at Button
            CollectionViewSource.GetDefaultView(Styles).MoveCurrentTo(Styles.Single(s => s.ElementType == typeof(Button)));
        }

        public List<StyleModel> Styles { get; private set; }

        private static IEnumerable<Type> GetFrameworkElementTypesFromAssembly(Assembly assembly)
        {
            // Returns all types in the specified assembly that are non-abstract,
            // and non-generic, derive from FrameworkElement, and have a default constructor

            foreach (var type in assembly.GetTypes())
            {
                if (// type.IsPublic && // maybe we wanna peek at nonpublic ones?
                    !type.IsAbstract &&
                    !type.ContainsGenericParameters &&
                    typeof(FrameworkElement).IsAssignableFrom(type) &&
                    (type.GetConstructor(Type.EmptyTypes) != null ||
                        type.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null) != null // allow a nonpublic ctor
                    ))
                {
                    yield return type;
                }
            }
        }

        private static IEnumerable<StyleModel> GetStyles(Assembly assembly)
        {
            return GetFrameworkElementTypesFromAssembly(assembly)
                .OrderBy(type => type.Name, StringComparer.Ordinal)
                .SelectMany(GetStyles);
        }

        private static IEnumerable<StyleModel> GetStyles(Type type)
        {
            // make an instance of the type and get its default style key
            if (type.GetConstructor(Type.EmptyTypes) != null)
            {
                var element = (FrameworkElement)Activator.CreateInstance(type, false);
                var defaultStyleKey = element.GetValue(DefaultStyleKeyProperty);

                yield return new StyleModel(
                    DisplayName: type.Name,
                    ResourceKey: defaultStyleKey,
                    ElementType: type);

                foreach (var styleModel in GetStylesFromStaticProperties(element))
                {
                    yield return styleModel;
                }
            }
        }

        private static IEnumerable<StyleModel> GetStylesFromStaticProperties(FrameworkElement element)
        {
            var properties = element.GetType()
                .GetProperties(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public)
                .Where(p => p.Name.EndsWith("StyleKey") && p.PropertyType == typeof(ResourceKey));

            foreach (var property in properties)
            {
                var elementType = element.GetType();
                var resourceKey = property.GetValue(element);

                yield return new StyleModel(
                    DisplayName: $"{elementType.Name}.{property.Name}",
                    ResourceKey: resourceKey,
                    ElementType: elementType);
            }
        }

        private void OnLoadClick(object sender, RoutedEventArgs e)
        {
            // create the file open dialog
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                CheckFileExists = true,
                Multiselect = false,
                Filter = "Assemblies (*.exe;*.dll)|*.exe;*.dll"
            };

            if (openFileDialog.ShowDialog(this) != true)
                return;

            try
            {
                AsmName.Text = openFileDialog.FileName;
                var styles = GetStyles(Assembly.LoadFile(openFileDialog.FileName)).ToList();
                if (styles.Count == 0)
                {
                    MessageBox.Show("Assembly does not contain any compatible types.");
                }
                else
                {
                    Styles = styles;
                    OnPropertyChanged(nameof(Styles));
                }
            }
            catch
            {
                MessageBox.Show("Error loading assembly.");
            }
        }

        private void ShowStyle(object sender, SelectionChangedEventArgs e)
        {
            if (styleTextBox == null) return;

            // see which type is selected
            if (typeComboBox.SelectedValue is StyleModel style)
            {
                var success = TrySerializeStyle(style.ResourceKey, out var serializedStyle);

                var styleXml = CleanupStyle(serializedStyle);

                // show the style in a document viewer
                styleTextBox.Document = CreateFlowDocument(success, styleXml.ToString());
            }
        }

        /// <summary>
        /// Serializes a style using XamlWriter.
        /// </summary>
        /// <param name="resourceKey"></param>
        /// <param name="serializedStyle"></param>
        /// <returns></returns>
        private static bool TrySerializeStyle(object resourceKey, out string serializedStyle)
        {
            var success = false;
            serializedStyle = "[Style not found]";

            if (resourceKey != null)
            {
                // try to get the default style for the type
                if (Application.Current.TryFindResource(resourceKey) is Style style)
                {
                    // try to serialize the style
                    try
                    {
                        var stringWriter = new StringWriter();
                        var xmlTextWriter = new XmlTextWriter(stringWriter) { Formatting = Formatting.Indented };
                        System.Windows.Markup.XamlWriter.Save(style, xmlTextWriter);
                        serializedStyle = stringWriter.ToString();

                        success = true;
                    }
                    catch (Exception exception)
                    {
                        serializedStyle = "[Exception thrown while serializing style]" +
                            Environment.NewLine + Environment.NewLine + exception;
                    }
                }
            }
            return success;
        }

        /// <summary>
        /// Creates a FlowDocument from the serialized XAML with simple syntax coloring.
        /// </summary>
        /// <param name="success"></param>
        /// <param name="serializedStyle"></param>
        /// <returns></returns>
        private FlowDocument CreateFlowDocument(bool success, string serializedStyle)
        {
            var document = new FlowDocument();
            if (success)
            {
                using (var reader = new XmlTextReader(serializedStyle, XmlNodeType.Document, null))
                {
                    var indent = 0;
                    var paragraph = new Paragraph();
                    while (reader.Read())
                    {
                        if (reader.IsStartElement()) // opening tag, e.g. <Button
                        {
                            string elementName = reader.Name;
                            // indentation
                            paragraph.AddRun(_textStyle, new string(' ', indent * 4));

                            paragraph.AddRun(_bracketStyle, "<");
                            paragraph.AddRun(_elementStyle, elementName);
                            if (reader.HasAttributes)
                            {
                                // write tag attributes
                                while (reader.MoveToNextAttribute())
                                {
                                    paragraph.AddRun(_attributeStyle, " " + reader.Name);
                                    paragraph.AddRun(_bracketStyle, "=");
                                    paragraph.AddRun(_quotesStyle, "\"");
                                    if (reader.Name == "TargetType") // target type fix - should use the Type MarkupExtension
                                    {
                                        paragraph.AddRun(_textStyle, "{x:Type " + reader.Value + "}");
                                    }
                                    else if (reader.Name == "Margin" || reader.Name == "Padding")
                                    {
                                        paragraph.AddRun(_textStyle, SimplifyThickness(reader.Value));
                                    }
                                    else
                                    {
                                        paragraph.AddRun(_textStyle, reader.Value);
                                    }

                                    paragraph.AddRun(_quotesStyle, "\"");
                                    paragraph.AddLineBreak();
                                    paragraph.AddRun(_textStyle, new string(' ', indent * 4 + elementName.Length + 1));
                                }
                                paragraph.RemoveLastLineBreak();
                                reader.MoveToElement();
                            }
                            if (reader.IsEmptyElement) // empty tag, e.g. <Button />
                            {
                                paragraph.AddRun(_bracketStyle, " />");
                                paragraph.AddLineBreak();
                                --indent;
                            }
                            else // non-empty tag, e.g. <Button>
                            {
                                paragraph.AddRun(_bracketStyle, ">");
                                paragraph.AddLineBreak();
                            }

                            ++indent;
                        }
                        else // closing tag, e.g. </Button>
                        {
                            --indent;

                            // indentation
                            paragraph.AddRun(_textStyle, new string(' ', indent * 4));

                            // text content of a tag, e.g. the text "Do This" in <Button>Do This</Button>
                            if (reader.NodeType == XmlNodeType.Text)
                            {
                                var value = reader.ReadContentAsString();
                                if (reader.Name == "Thickness")
                                    value = SimplifyThickness(value);
                                paragraph.AddRun(_textStyle, value);
                            }

                            paragraph.AddRun(_bracketStyle, "</");
                            paragraph.AddRun(_elementStyle, reader.Name);
                            paragraph.AddRun(_bracketStyle, ">");
                            paragraph.AddLineBreak();
                        }
                    }
                    document.Blocks.Add(paragraph);
                }
            }
            else // no style found
            {
                document.Blocks.Add(new Paragraph(new Run(serializedStyle)) { TextAlignment = TextAlignment.Left });
            }
            return document;
        }

        private static string SimplifyThickness(string s)
        {
            var four = Regex.Match(s, @"(-?[\d+]),\1,\1,\1");
            if (four.Success)
                return four.Groups[1].Value;
            var two = Regex.Match(s, @"(-?[\d+]),(-?[\d+]),\1,\2");
            if (two.Success)
                return $"{two.Groups[1].Value},{two.Groups[2].Value}";
            return s;
        }

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        private static readonly XNamespace xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
        private static readonly XNamespace xmlns_s = "clr-namespace:System;assembly=mscorlib";
        private static readonly XNamespace xmlns_x = "http://schemas.microsoft.com/winfx/2006/xaml";

        private static XDocument CleanupStyle(string serializedStyle)
        {
            XDocument styleXml = XDocument.Parse(serializedStyle);

            RemoveEmptyResources(styleXml);
            SimplifyStyleSetterValues(styleXml);

            return styleXml;
        }

        private static void RemoveEmptyResources(XDocument styleXml)
        {
            foreach (var elt in styleXml.Descendants())
            {
                var localName = elt.Name.LocalName;

                var eltResources = elt.Element(xmlns + $"{localName}.Resources");
                var eltResourceDictionary = eltResources?.Element(xmlns + "ResourceDictionary");

                if (eltResourceDictionary?.IsEmpty ?? false)
                    eltResources.Remove();
            }
        }

        private static void SimplifyStyleSetterValues(XDocument styleXml)
        {
            foreach (var elt in styleXml.Descendants())
            {
                var localName = elt.Name.LocalName;

                var eltValueNode = elt.Element(xmlns + $"{localName}.Value");
                var eltValue = eltValueNode?.Elements().SingleOrDefault();

                switch (eltValue?.Name)
                {
                    case { } name when name == xmlns + "SolidColorBrush":
                        elt.SetAttributeValue("Value", eltValue.Value);
                        eltValueNode.Remove();
                        break;
                    case { } name when name == xmlns + "DynamicResource":
                        elt.SetAttributeValue("Value", $"{{DynamicResource {eltValue.Attribute("ResourceKey")?.Value}}}");
                        eltValueNode.Remove();
                        break;
                    case { } name when name == xmlns + "StaticResource":
                        elt.SetAttributeValue("Value", $"{{StaticResource {eltValue.Attribute("ResourceKey")?.Value}}}");
                        eltValueNode.Remove();
                        break;
                    case { } name when name.Namespace == xmlns_s:
                        elt.SetAttributeValue("Value", eltValue.Value);
                        eltValueNode.Remove();
                        break;
                    case { } name when name == xmlns + "Thickness":
                        elt.SetAttributeValue("Value", SimplifyThickness(eltValue.Value));
                        eltValueNode.Remove();
                        break;
                    case { } name when name == xmlns_x + "Static":
                        {
                            var value = eltValue.Attribute("Member")?.Value;
                            value = value?.Split('.').Last();
                            if (value != null)
                            {
                                elt.SetAttributeValue("Value", value);
                                eltValueNode.Remove();
                            }
                            break;
                        }
                }

            }
        }
    }
}
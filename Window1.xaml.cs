using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Navigation;
using System.Xml;
using System.Windows.Documents;
using System.Reflection;

namespace StyleSnooper
{
    public partial class Window1
    {
        Style bracketStyle, elementStyle, quotesStyle, textStyle, attributeStyle, commentStyle;
        Microsoft.Win32.OpenFileDialog openFileDialog;

        public Window1()
        {
            this.InitializeComponent();

            // get syntax coloring styles
            bracketStyle = Resources["BracketStyle"] as Style;
            elementStyle = Resources["ElementStyle"] as Style;
            quotesStyle = Resources["QuotesStyle"] as Style;
            textStyle = Resources["TextStyle"] as Style;
            attributeStyle = Resources["AttributeStyle"] as Style;
            commentStyle = Resources["CommentStyle"] as Style;

            // start out by looking at Button
            CollectionViewSource.GetDefaultView(this.ElementTypes).MoveCurrentTo(typeof(Button));

            // create the file open dialog
            openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Assemblies (*.exe;*.dll)|*.exe;*.dll";
        }

        void SandAndFire(object sender, EventArgs e)
        {
            GlassHelper.TryExtendGlassFrame(this, styleTextBox.Margin);
        }

        private Type[] elementTypes;
        public Type[] ElementTypes
        {
            get
            {
                if (elementTypes == null)
                {
                    elementTypes = GetFrameworkElemenetTypesFromAssembly(typeof(FrameworkElement).Assembly);
                }

                return elementTypes;
            }
        }

        private Type[] GetFrameworkElemenetTypesFromAssembly(Assembly asm)
        {
            // create a list of all types in PresentationFramework that are non-abstract,
            // and non-generic, derive from FrameworkElement, and have a default constructor
            List<Type> typeList = new List<Type>();
            foreach (Type type in asm.GetTypes())
            {
                if (// type.IsPublic && // maybe we wanna peek at nonpublic ones?
                    !type.IsAbstract &&
                    !type.ContainsGenericParameters &&
                    typeof(FrameworkElement).IsAssignableFrom(type) &&
                    (type.GetConstructor(Type.EmptyTypes) != null ||
                        type.GetConstructor(BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null) != null // allow a nonpublic ctor
                    ))
                {
                    typeList.Add(type);
                }
            }

            // sort the types by name
            typeList.Sort(delegate(Type typeA, Type typeB)
            {
                return String.CompareOrdinal(typeA.Name, typeB.Name);
            });
            
            return typeList.ToArray();
        }

        void OnLoadClick(object sender, RoutedEventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == true)
            {
                try
                {
                    AsmName.Text = openFileDialog.FileName;
                    Type[] types = GetFrameworkElemenetTypesFromAssembly(Assembly.LoadFile(openFileDialog.FileName));
                    if (types.Length == 0)
                    {
                        MessageBox.Show("Assembly does not contain any compatible types.");
                    }
                    else
                    {
                        elementTypes = types;

                        BindingExpression exp = BindingOperations.GetBindingExpression(typeComboBox, ComboBox.ItemsSourceProperty);
                        exp.UpdateTarget();
                    }
                }
                catch
                {
                    MessageBox.Show("Error loading assembly.");
                }
            }
        }

        private void ShowStyle(object sender, SelectionChangedEventArgs e)
        {
            if (styleTextBox == null) return;

            // see which type is selected
            Type type = this.typeComboBox.SelectedValue as Type;
            if (type != null)
            {
                string serializedStyle;
                bool success = TrySerializeStyle(type, out serializedStyle);

                // show the style in a document viewer
                this.styleTextBox.Document = CreateFlowDocument(success, serializedStyle);
            }
        }

        /// <summary>
        /// Serializes a style using XamlWriter.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="serializedStyle"></param>
        /// <returns></returns>
        private static bool TrySerializeStyle(Type type, out string serializedStyle)
        {
            bool success = false;
            serializedStyle = "[Style not found]";
            
            // make an instance of the type and get its default style key
            bool nonPublic = type.GetConstructor(Type.EmptyTypes) == null;
            FrameworkElement element = (FrameworkElement)Activator.CreateInstance(type, nonPublic);

            object defaultStyleKey = element.GetValue(FrameworkElement.DefaultStyleKeyProperty);
            
            if (defaultStyleKey != null)
            {
                // try to get the default style for the type
                Style style = Application.Current.TryFindResource(defaultStyleKey) as Style;
                if (style != null)
                {
                    // try to serialize the style
                    try
                    {
                        StringWriter stringWriter = new StringWriter();
                        XmlTextWriter xmlTextWriter = new XmlTextWriter(stringWriter);
                        xmlTextWriter.Formatting = Formatting.Indented;
                        System.Windows.Markup.XamlWriter.Save(style, xmlTextWriter);
                        serializedStyle = stringWriter.ToString();

                        success = true;
                    }
                    catch (Exception exception)
                    {
                        serializedStyle = "[Exception thrown while serializing style]" +
                            Environment.NewLine + Environment.NewLine + exception.ToString();
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
            FlowDocument document = new FlowDocument();
            if (success)
            {
                using (XmlTextReader reader = new XmlTextReader(serializedStyle, XmlNodeType.Document, null))
                {
                    int indent = 0;
                    Paragraph paragraph = new Paragraph();
                    while (reader.Read())
                    {
                        if (reader.IsStartElement()) // opening tag, e.g. <Button
                        {
                            // indentation
                            AddRun(paragraph, textStyle, new string(' ', indent * 4));

                            AddRun(paragraph, bracketStyle, "<");
                            AddRun(paragraph, elementStyle, reader.Name);
                            if (reader.HasAttributes)
                            {
                                // write tag attributes
                                while (reader.MoveToNextAttribute())
                                {
                                    AddRun(paragraph, attributeStyle, " " + reader.Name);
                                    AddRun(paragraph, bracketStyle, "=");
                                    AddRun(paragraph, quotesStyle, "\"");
                                    if (reader.Name == "TargetType") // target type fix - should use the Type MarkupExtension
                                    {
                                        AddRun(paragraph, textStyle, "{x:Type " + reader.Value + "}");
                                    }
                                    else
                                    {
                                        AddRun(paragraph, textStyle, reader.Value);
                                    }
                                    AddRun(paragraph, quotesStyle, "\"");
                                }
                                reader.MoveToElement();
                            }
                            if (reader.IsEmptyElement) // empty tag, e.g. <Button />
                            {
                                AddRun(paragraph, bracketStyle, " />");
                                paragraph.Inlines.Add(new LineBreak());
                                --indent;
                            }
                            else // non-empty tag, e.g. <Button>
                            {
                                AddRun(paragraph, bracketStyle, ">");
                                paragraph.Inlines.Add(new LineBreak());
                            }

                            ++indent;
                        }
                        else // closing tag, e.g. </Button>
                        {
                            // indentation
                            AddRun(paragraph, textStyle, new string(' ', indent * 4));

                            // text content of a tag, e.g. the text "Do This" in <Button>Do This</Button>
                            if (reader.NodeType == XmlNodeType.Text)
                            {
                                AddRun(paragraph, textStyle, reader.ReadContentAsString());
                            }

                            AddRun(paragraph, bracketStyle, "</");
                            AddRun(paragraph, elementStyle, reader.Name);
                            AddRun(paragraph, bracketStyle, ">");
                            paragraph.Inlines.Add(new LineBreak());

                            --indent;
                        }
                    }
                    document.Blocks.Add(paragraph);
                }
            }
            else // no style found
            {
                Paragraph par = new Paragraph(new Run(serializedStyle));
                par.TextAlignment = TextAlignment.Left;
                document.Blocks.Add(par);
            }
            return document;
        }

        /// <summary>
        /// Adds a span with the specified text and style.
        /// </summary>
        /// <param name="par"></param>
        /// <param name="style"></param>
        /// <param name="s"></param>
        private static void AddRun(Paragraph par, Style style, string s)
        {
            Run run = new Run(s);
            run.Style = style;
            par.Inlines.Add(run);
        }
    }
}
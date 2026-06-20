using System;

namespace QicBoqMapper
{
    public class QicSettings
    {
        public string SeparatorStyle { get; set; } = "Dash"; // Dash, Dot, Underscore
        public string MatchingMethod { get; set; } = "Category + Family + Type"; // Element ID, Category + Family + Type
        public string SelectionScope { get; set; } = "Entire Model"; // Current Selection, Active View, Entire Model
        public bool CaseInsensitive { get; set; } = true;
    }
}

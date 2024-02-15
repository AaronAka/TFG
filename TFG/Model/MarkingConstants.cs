namespace TFG.Model
{
    public enum ParagraphType
    {
        Empty,
        Doctitle,
        Authors,
        Sectitle,
        Subsection,
        Regular,
        Table,
        Keywords,
        Sources,
        Bibliography
    }

    public static class MarkingConstants
    {
        public static string? FILEPATH;
        public static string? FILENAME;

        #region Tags
        public const string DOCTITLE_OPEN = "[doctitle]";
        public const string DOCTITLE_CLOSE = "[/doctitle]";
        public const string BIBLIOGRAPHY_REF_OPEN = "[ref id=\"r{0}\" reftype =\"journal\"][authors role=\"nd\"]";
        public const string BIBLIOGRAPHY_REF_CLOSE = "[/ref]";
        public const string BIBLIOGRAPHY_PAUTHOR_OPEN = " [pauthor][surname]{0}[/surname],";
        public const string BIBLIOGRAPHY_PAUTHOR_CLOSE = " [fname]{0}[/fname][/pauthor],";
        public const string BIBLIOGRAPHY_PAUTHOR_CLOSE_NO_COMMA = " [fname]{0}[/fname][/pauthor]";
        public const string BIBLIOGRAPHY_AUTHORS_CLOSE = "[/authors]";
        public const string BIBLIOGRAPHY_DATE = "([date dateiso=\"" + "{0}" + "0000" + "\" specyear=\"{0}\"]{0}[/date]).";
        public const string BIBLIOGRAPHY_ARTTITLE = " [arttitle]{0}.[/arttitle]";
        public const string BIBLIOGRAPHY_SOURCE = "[source]{0}[/source],";
        public const string BIBLIOGRAPHY_VOLID = "[volid]{0}[/volid]";
        public const string BIBLIOGRAPHY_ISSUENO = "([issueno]{0}[/issueno]),";
        public const string BIBLIOGRAPHY_PAGES = "[pages]{0}[/pages]";
        public const string BIBLIOGRAPHY_PUBID = "[pubid]{0}[/pubid]";
        public const string AUTHORS_OPEN = "[author role=\"nd\" rid=\"aff1\" corresp=\"n\" deceased=\"n\" eqcontr=\"nd\"][surname]{0}[/surname]";
        public const string AUTHORS_CLOSE = ", [fname]{0}.[/fname][/author]";
        public const string INITIAL_SECTYPE_OPEN = "[xmlbody]\r[sec sec-type=\"{0}\"]";
        public const string SECTYPE_OPEN = "[sec sec-type=\"{0}\"]";
        public const string SECTITLE_OPEN = "[sectitle]";
        public const string SECTITLE_CLOSE = "[/sectitle]";
        public const string SECTITLE = "[sectitle]{0}[/sectitle]";
        public const string REGULAR_PARAGRAPH_OPEN = "[p]";
        public const string REGULAR_PARAGRAPH_CLOSE = "[/p]";
        public const string REGULAR_PARAGRAPH = "[p]{0}[/p]";
        public const string FIGGRP_OPEN = "[figgrp id ={0}]";
        public const string LABEL_OPEN = "[label]";
        public const string LABEL_CLOSE = "[/label]";
        public const string CAPTION_OPEN = "[caption]";
        public const string CAPTION_CLOSE = "[/caption]";
        public const string KWDGRP_OPEN = "[kwdgrp]";
        public const string KWDGRP_CLOSE = "[/kwdgrp]";
        public const string KWD_OPEN = "[kwd]";
        public const string KWD_CLOSE = "[/kwd]";
        public const string NORMAFF_OPEN = "[normaff id =\"aff{0}\" ncountry=\"{1}\" ]";
        public const string NORMAFF_CLOSE = "[/normaff]";
        public const string SUP_OPEN = "[sup]";
        public const string SUP_CLOSE = "[/sup]";
        public const string ORGDIVE_OPEN = "[orgdiv{0}]";
        public const string ORGDIVE_CLOSE = "[/orgdiv{0}]";
        public const string CITY_OPEN = "[city]";
        public const string CITY_CLOSE = "[/city]";
        public const string COUNTRY_OPEN = "[country]";
        public const string COUNTRY_CLOSE = "[/country]";

        #endregion

        #region Literals
        public const string OPEN_PARENTHESIS = "(";
        public const string CLOSE_PARENTHESIS = ")";
        public const string DOT = ".";
        public const string COLON = ":";
        public const string SEMICOLON = ";";
        public const string COMMA = ",";
        #endregion
    }
}
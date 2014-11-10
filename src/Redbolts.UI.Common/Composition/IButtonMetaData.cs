namespace Redbolts.UI.Common.Composition
{
    public interface IButtonMetaData : IRibbonItemMetaData
    {
        string Text { get; }
        string Image { get; }
        string LargeImage { get; }
    }
}
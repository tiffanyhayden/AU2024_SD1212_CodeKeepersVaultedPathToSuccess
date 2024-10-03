using Inventor;

namespace ExternalRuleManager
{
    /// <summary>
    /// Provides methods to convert between <see cref="Image"/> and <see cref="IPictureDisp"/> objects.
    /// </summary>
    public static class ImageConverter
    {
        /// <summary>
        /// Converts a <see cref="Image"/> to an <see cref="IPictureDisp"/> object.
        /// </summary>
        /// <param name="image">The <see cref="Image"/> to convert.</param>
        /// <returns>
        /// The converted <see cref="IPictureDisp"/> object, or <c>null</c> if the conversion fails.
        /// </returns>
        /// <remarks>
        /// This method uses <see cref="ImageConverterUtilities.ToIPictureDisp"/> internally.
        /// </remarks>
        public static IPictureDisp ConvertImageToIPictureDisp(Image image)
        {
            try
            {
                return ImageConverterUtilities.ToIPictureDisp(image);
            }
            catch (Exception)
            {
                // Handle or log exceptions if needed.
                return null;
            }
        }

        /// <summary>
        /// Converts an <see cref="IPictureDisp"/> object to an <see cref="Image"/>.
        /// </summary>
        /// <param name="iPict">The <see cref="IPictureDisp"/> object to convert.</param>
        /// <returns>
        /// The converted <see cref="Image"/> object, or <c>null</c> if the conversion fails.
        /// </returns>
        /// <remarks>
        /// This method uses <see cref="ImageConverterUtilities.ToImage"/> internally.
        /// </remarks>
        public static Image ConvertIPictureDispToImage(IPictureDisp iPict)
        {
            try
            {
                return ImageConverterUtilities.ToImage(iPict);
            }
            catch (Exception)
            {
                // Handle or log exceptions if needed.
                return null;
            }
        }
    }
}

internal static class ImageConverterUtilities
{
    /// <summary>
    /// Contains utility methods for converting between <see cref="Image"/> and <see cref="IPictureDisp"/> objects.
    /// These methods leverage the <see cref="AxHost"/> class for the conversion process.
    /// </summary>
    private class AxHostConverter : AxHost
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="AxHostConverter"/> class.
        /// </summary>
        private AxHostConverter() : base(string.Empty) { }

        /// <summary>
        /// Converts a given <see cref="Image"/> to an <see cref="IPictureDisp"/> object.
        /// </summary>
        /// <param name="image">The <see cref="Image"/> to be converted.</param>
        /// <returns>An <see cref="IPictureDisp"/> object representing the converted image.</returns>
        public static IPictureDisp GetIPictureDisp(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        /// <summary>
        /// Converts an <see cref="IPictureDisp"/> object to a <see cref="Image"/>.
        /// </summary>
        /// <param name="pictureDisp">The <see cref="IPictureDisp"/> object to be converted.</param>
        /// <returns>An <see cref="Image"/> object representing the converted <see cref="IPictureDisp"/>.</returns>
        public static Image GetImageFromIPictureDisp(IPictureDisp pictureDisp)
        {
            return GetPictureFromIPictureDisp(pictureDisp);
        }
    }

    /// <summary>
    /// Converts an <see cref="Image"/> to an <see cref="IPictureDisp"/> object using <see cref="AxHost"/>.
    /// </summary>
    /// <param name="image">The <see cref="Image"/> to be converted.</param>
    /// <returns>The converted <see cref="IPictureDisp"/> object.</returns>
    public static IPictureDisp ToIPictureDisp(Image image)
    {
        return AxHostConverter.GetIPictureDisp(image);
    }

    /// <summary>
    /// Converts an <see cref="IPictureDisp"/> object to an <see cref="Image"/> using <see cref="AxHost"/>.
    /// </summary>
    /// <param name="picturedDisp">The <see cref="IPictureDisp"/> object to convert.</param>
    /// <returns>The converted <see cref="Image"/> object.</returns>
    public static Image ToImage(IPictureDisp picturedDisp)
    {
        return AxHostConverter.GetImageFromIPictureDisp(picturedDisp);
    }
}
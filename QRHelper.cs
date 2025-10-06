using QRCoder;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace PBQRCodeLib
{
    [ComVisible(true)]
    [Guid("12345678-1234-1234-1234-123456789012")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class QRHelper
    {
    
        public int GenerateQRCode(string text, string qrPath, string logoPath)
        {
            try
            {
                if (string.IsNullOrEmpty(text))
                    return 0;

                if (string.IsNullOrEmpty(qrPath))
                    return 0;

                using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
                {
                    QRCodeData qrCodeData = qrGenerator.CreateQrCode(text, QRCodeGenerator.ECCLevel.Q);
                    using (QRCode qrCode = new QRCode(qrCodeData))
                    using (Bitmap qrImage = qrCode.GetGraphic(20, Color.Black, Color.White, true))
                    {
                        if (!string.IsNullOrEmpty(logoPath) && File.Exists(logoPath))
                        {
                            using (Graphics g = Graphics.FromImage(qrImage))
                            using (Image logo = Image.FromFile(logoPath))
                            {
                                // Resize logo (20% of QR code size)
                                int logoSize = qrImage.Width / 5;
                                int x = (qrImage.Width - logoSize) / 2;
                                int y = (qrImage.Height - logoSize) / 2;
                                
                                // Use high-quality rendering
                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                g.DrawImage(logo, new Rectangle(x, y, logoSize, logoSize));
                            }
                        }

                        // Ensure directory exists
                        string directory = Path.GetDirectoryName(qrPath);
                        if (!Directory.Exists(directory))
                            Directory.CreateDirectory(directory);

                        qrImage.Save(qrPath, ImageFormat.Png);
                        return 1; // Success
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"QR Generation Error: {ex.Message}");
                return 0; // Failure
            }
        }

        /// <summary>
        /// Generate simple QR Code without logo
        /// </summary>
        public int GenerateSimpleQRCode(string text, string qrPath)
        {
            return GenerateQRCode(text, qrPath, "");
        }
    }
}
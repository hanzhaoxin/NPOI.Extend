/*
 类：SheetExtend
 描述：Sheet扩展方法
 编 码 人：韩兆新 日期：2015年04月11日
 修改记录：
    1: 修改人：韩兆新  日期：2015年11月21日
       修改内容：添加了移动图片的方法
   

*/
using System;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Extend
{
    public static partial class SheetExtend
    {
        #region 1.0 添加图片

        /// <summary>
        ///     添加图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="picInfo"></param>
        public static void AddPicture(this ISheet sheet, PictureInfo picInfo)
        {
            int pictureIdx = sheet.Workbook.AddPicture(picInfo.PictureData, PictureType.PNG);
            IClientAnchor anchor = sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = picInfo.MinCol;
            anchor.Col2 = picInfo.MaxCol;
            anchor.Row1 = picInfo.MinRow;
            anchor.Row2 = picInfo.MaxRow;
            anchor.Dx1 = picInfo.PicturesStyle.AnchorDx1;
            anchor.Dx2 = picInfo.PicturesStyle.AnchorDx2;
            anchor.Dy1 = picInfo.PicturesStyle.AnchorDy1;
            anchor.Dy2 = picInfo.PicturesStyle.AnchorDy2;
            anchor.AnchorType = AnchorType.MoveDontResize;
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IPicture pic = drawing.CreatePicture(anchor, pictureIdx);
            if (sheet is HSSFSheet)
            {
                var shape = (HSSFShape) pic;
                shape.FillColor = picInfo.PicturesStyle.FillColor;
                shape.IsNoFill = picInfo.PicturesStyle.IsNoFill;
                shape.LineStyle = picInfo.PicturesStyle.LineStyle;
                shape.LineStyleColor = picInfo.PicturesStyle.LineStyleColor;
                shape.LineWidth = (int) picInfo.PicturesStyle.LineWidth;
            }
            else if (sheet is XSSFSheet)
            {
                var shape = (XSSFShape) pic;
                shape.FillColor = picInfo.PicturesStyle.FillColor;
                shape.IsNoFill = picInfo.PicturesStyle.IsNoFill;
                shape.LineStyle = picInfo.PicturesStyle.LineStyle;
                //shape.LineStyleColor = picInfo.PicturesStyle.LineStyleColor;
                shape.LineWidth = picInfo.PicturesStyle.LineWidth;
            }
        }

        #endregion

        #region 1.1 获取图片信息

        /// <summary>
        ///     获取sheet中包含图片的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public static List<PictureInfo> GetAllPictureInfos(this ISheet sheet)
        {
            return sheet.GetAllPictureInfos(null, null, null, null);
        }

        /// <summary>
        ///     获取sheet中指定区域包含图片的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        public static List<PictureInfo> GetAllPictureInfos(this ISheet sheet, int? minRow, int? maxRow, int? minCol,
            int? maxCol, bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                return GetAllPictureInfos((HSSFSheet) sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            if (sheet is XSSFSheet)
            {
                return GetAllPictureInfos((XSSFSheet) sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            throw new Exception("未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！");
        }

        #endregion

        #region 1.3 删除图片

        /// <summary>
        ///     清除sheet中的图片
        /// </summary>
        /// <param name="sheet"></param>
        public static void RemovePictures(this ISheet sheet)
        {
            sheet.RemovePictures(null, null, null, null);
        }

        /// <summary>
        ///     清除sheet中指定区域的图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        public static void RemovePictures(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
            bool onlyInternal = true)
        {
            if (sheet is HSSFSheet)
            {
                RemovePictures((HSSFSheet) sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else if (sheet is XSSFSheet)
            {
                RemovePictures((XSSFSheet) sheet, minRow, maxRow, minCol, maxCol, onlyInternal);
            }
            else
            {
                throw new Exception("未处理类型，没有为该类型添加：GetAllPicturesInfos()扩展方法！");
            }
        }

        #endregion

        #region 1.4 移动图片

        /// <summary>
        ///     移动sheet中所有图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="moveRowCount"></param>
        /// <param name="moveColCount"></param>
        public static void MovePictures(this ISheet sheet, int moveRowCount = 0, int moveColCount = 0)
        {
            sheet.MovePictures(null, null, null, null, true, moveRowCount, moveColCount);
        }

        /// <summary>
        ///     移动指定区域内的图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <param name="moveRowCount"></param>
        /// <param name="moveColCount"></param>
        public static void MovePictures(this ISheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
            bool onlyInternal = true,
            int moveRowCount = 0, int moveColCount = 0)
        {
            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                IList<HSSFShape> shapeList = shapeContainer.Children;
                foreach (HSSFShape shape in shapeList)
                {
                    var anchor = shape.Anchor as IClientAnchor;
                    if (null != anchor)
                    {
                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1,
                            anchor.Col2, onlyInternal))
                        {
                            anchor.Row1 += moveRowCount;
                            anchor.Row2 += moveRowCount;
                            anchor.Col1 += moveColCount;
                            anchor.Col2 += moveColCount;
                        }
                    }
                }
            }
        }

        #endregion

        #region 3.0 实现（获取图片信息）

        /// <summary>
        ///     HSSFSheet获取指定区域包含图片的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static List<PictureInfo> GetAllPictureInfos(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol,
            int? maxCol, bool onlyInternal)
        {
            var picturesInfoList = new List<PictureInfo>();

            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                IList<HSSFShape> shapeList = shapeContainer.Children;
                foreach (HSSFShape shape in shapeList)
                {
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture) shape;
                        var anchor = (HSSFClientAnchor) shape.Anchor;
                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1,
                            anchor.Col2, onlyInternal))
                        {
                            var picStyle = new PictureStyle
                            {
                                AnchorDx1 = anchor.Dx1,
                                AnchorDx2 = anchor.Dx2,
                                AnchorDy1 = anchor.Dy1,
                                AnchorDy2 = anchor.Dy2,
                                IsNoFill = picture.IsNoFill,
                                LineStyle = picture.LineStyle,
                                LineStyleColor = picture.LineStyleColor,
                                LineWidth = picture.LineWidth,
                                FillColor = picture.FillColor
                            };
                            picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2,
                                picture.PictureData.Data, picStyle));
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        /// <summary>
        ///     XSSFSheet获取指定区域包含图片的信息列表
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        /// <returns></returns>
        private static List<PictureInfo> GetAllPictureInfos(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol,
            int? maxCol, bool onlyInternal)
        {
            var picturesInfoList = new List<PictureInfo>();

            List<POIXMLDocumentPart> documentPartList = sheet.GetRelations();
            foreach (POIXMLDocumentPart documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing) documentPart;
                    List<XSSFShape> shapeList = drawing.GetShapes();
                    foreach (XSSFShape shape in shapeList)
                    {
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture) shape;
                            IClientAnchor anchor = picture.GetPreferredSize();

                            if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2,
                                anchor.Col1, anchor.Col2, onlyInternal))
                            {
                                var picStyle = new PictureStyle
                                {
                                    AnchorDx1 = anchor.Dx1,
                                    AnchorDx2 = anchor.Dx2,
                                    AnchorDy1 = anchor.Dy1,
                                    AnchorDy2 = anchor.Dy2,
                                    IsNoFill = picture.IsNoFill,
                                    LineStyle = picture.LineStyle,
                                    LineStyleColor = picture.LineStyleColor,
                                    LineWidth = picture.LineWidth,
                                    FillColor = picture.FillColor
                                };
                                picturesInfoList.Add(new PictureInfo(anchor.Row1, anchor.Row2, anchor.Col1, anchor.Col2,
                                    picture.PictureData.Data, picStyle));
                            }
                        }
                    }
                }
            }

            return picturesInfoList;
        }

        #endregion

        #region 3.1 实现（删除图片）

        /// <summary>
        ///     HSSFSheet清除指定区域的图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        private static void RemovePictures(HSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
            bool onlyInternal)
        {
            var shapeContainer = sheet.DrawingPatriarch as HSSFShapeContainer;
            if (null != shapeContainer)
            {
                IList<HSSFShape> shapeList = shapeContainer.Children;
                for (int i = 0; i < shapeList.Count; i++)
                {
                    HSSFShape shape = shapeList[i];
                    if (shape is HSSFPicture && shape.Anchor is HSSFClientAnchor)
                    {
                        var picture = (HSSFPicture) shape;
                        var anchor = (HSSFClientAnchor) shape.Anchor;
                        if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2, anchor.Col1,
                            anchor.Col2, onlyInternal))
                        {
                            shapeContainer.RemoveShape(shape);
                        }
                    }
                }
            }
        }

        /// <summary>
        ///     XSSFSheet清除指定区域的图片
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="minRow"></param>
        /// <param name="maxRow"></param>
        /// <param name="minCol"></param>
        /// <param name="maxCol"></param>
        /// <param name="onlyInternal"></param>
        private static void RemovePictures(XSSFSheet sheet, int? minRow, int? maxRow, int? minCol, int? maxCol,
            bool onlyInternal)
        {
            List<POIXMLDocumentPart> documentPartList = sheet.GetRelations();
            foreach (POIXMLDocumentPart documentPart in documentPartList)
            {
                if (documentPart is XSSFDrawing)
                {
                    var drawing = (XSSFDrawing) documentPart;
                    List<XSSFShape> shapeList = drawing.GetShapes();

                    for (int i = 0; i < shapeList.Count; i++)
                    {
                        XSSFShape shape = shapeList[i];
                        if (shape is XSSFPicture)
                        {
                            var picture = (XSSFPicture) shape;
                            IClientAnchor anchor = picture.GetPreferredSize();

                            if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, anchor.Row1, anchor.Row2,
                                anchor.Col1, anchor.Col2, onlyInternal))
                            {
                                throw new NotImplementedException("XSSFSheet未实现ClearPictures()方法！");
                            }
                        }
                    }
                }
            }
        }

        #endregion
    }
}
/**
 * Change the layouts of an existing section
 */
export enum configLayout {
    'SingleColumnSectionToolboxItem' = 'PropertyPaneChoiceGroup-10',
    'DoubleColumnSectionToolboxItem' = 'PropertyPaneChoiceGroup-5',
    'TripleColumnSectionToolboxItem' = 'PropertyPaneChoiceGroup-6',
    'ColumnRightTwoThirdsSectionToolboxItem' = 'PropertyPaneChoiceGroup-8',
    'ColumnLeftTwoThirdsSectionToolboxItem' = 'PropertyPaneChoiceGroup-7',
}

/**
 * Available layouts for a new section
 */
export type pageLayout =
    'SingleColumnSectionToolboxItem'
    | 'DoubleColumnSectionToolboxItem'
    | 'TripleColumnSectionToolboxItem'
    | 'ColumnRightTwoThirdsSectionToolboxItem'
    | 'ColumnLeftTwoThirdsSectionToolboxItem'
    | 'FullWidthSectionToolboxItem'
    | 'VerticalSectionToolboxItem';

/**
 * Available Background colors for a section
 */
export type background =
    'noneBackgroundColorButton'
    | 'neutralBackgroundColorButton'
    | 'softBackgroundColorButton'
    | 'strongBackgroundColorButton';
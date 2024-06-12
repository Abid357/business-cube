import java.awt.Component;

import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.table.TableCellRenderer;

public class WrapTextRenderer extends JTextArea implements TableCellRenderer {

private static final long serialVersionUID = 1L;

public WrapTextRenderer() {
setAlignmentY(BOTTOM_ALIGNMENT);
setLineWrap(true);
setWrapStyleWord(true);
setOpaque(true);
}

@Override
public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
    int row, int column) {
// TODO Auto-generated method stub
setText((String)value);//or something in value, like value.getNote()..
return this;
  } 
}
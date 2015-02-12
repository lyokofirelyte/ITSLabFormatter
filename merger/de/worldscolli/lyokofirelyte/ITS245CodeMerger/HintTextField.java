package de.worldscolli.lyokofirelyte.ITS245CodeMerger;

import java.awt.Color;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;

import javax.swing.BorderFactory;
import javax.swing.JTextField;

public class HintTextField extends JTextField implements FocusListener {
	
	private static final long serialVersionUID = -2303974084975946640L;
	
	private String hint;
	private Color colour;
	
	private boolean showing;
	
	public HintTextField(String hint){
		
		super(hint);
		colour = new Color(155, 155, 155);
		
		this.hint = hint;
		this.showing = true;
		
		super.setBorder(BorderFactory.createCompoundBorder(BorderFactory.createLineBorder(new Color(216, 216, 216)), BorderFactory.createEmptyBorder(2, 5, 2, 2)));
		super.addFocusListener(this);
		
		super.setBackground(Color.white);
		super.setForeground(colour);
		
	}
	
	public void focusGained(FocusEvent e){
		
		if (getText().isEmpty()){
			
			hideHint();
			
		}
		
	}
	
	public void focusLost(FocusEvent e){
		
		if (getText().isEmpty()){
			
			showHint();
			
		}
		
	}
	
	@Override
	public String getText(){
		
		return showing ? "" : super.getText();
		
	}
	
	@Override
	public void setText(String text){
		
		if (text.equals("")){
			
			showHint();
			
		} else {
			
			hideHint();
			super.setText(text);
			
		}
		
	}
	
	public void setHint(String hint){
		
		hideHint();
		this.hint = hint;
		showHint();
	}
	
	private void showHint(){
		
		super.setText(hint);
		super.setForeground(colour);
		
		showing = true;
		
	}
	
	private void hideHint(){
		
		super.setText("");
		super.setForeground(new Color(45, 45, 45));
		
		showing = false;
		
	}
	
}

import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { IFieldFormatting, IDateFormatRule } from '../PowerformWebPart';
import { LoggerService } from './LoggerService';

export interface IFormattingEditorProps {
    field: { key: string; text: string; type: string; choices?: string[] };
    config: IFieldFormatting;
    onSave: (config: IFieldFormatting) => void;
    onCancel: () => void;
}

export class FormattingEditor extends React.Component<IFormattingEditorProps, IFieldFormatting> {
    constructor(props: IFormattingEditorProps) {
        super(props);
        try {
            this.state = props.config || {
                type: props.field.type === 'DateTime' ? 'date' : 'choice',
                choiceConfig: {},
                dateRules: []
            };
        } catch (error:any) {
            void LoggerService.log(
                'FormattingEditor-constructor', 
                'High', 
                'Config', 
                error instanceof Error ? error.message : JSON.stringify(error)
            );
        }
    }

    private _saveChoiceColor = (choice: string, color: string): void => {
        try {
            const choiceConfig = { ...this.state.choiceConfig, [choice]: color };
            this.setState({ choiceConfig });
        } catch (error:any) {
            void LoggerService.log(
                'FormattingEditor-saveChoiceColor', 
                'Low', 
                'Config', 
                error instanceof Error ? error.message : JSON.stringify(error)
            );
        }
    };

    private _addDateRule = (): void => {
        try {
            const dateRules = [...(this.state.dateRules || [])];
            dateRules.push({ condition: 'past', days: 0, color: '#ffbb00' });
            this.setState({ dateRules });
        } catch (error:any) {
            void LoggerService.log(
                'FormattingEditor-addDateRule', 
                'Medium', 
                'Config', 
                error instanceof Error ? error.message : JSON.stringify(error)
            );
        }
    };

    public render(): React.ReactElement<IFormattingEditorProps> {
        const { field } = this.props;
        const isDate = field.type === 'DateTime';

        return (
            <div style={{ padding: '10px 0' }}>
                <h3 style={{ marginTop: 0, fontSize: '16px', fontWeight: 600, color: '#323130' }}>
                    Formatting: {field.text}
                </h3>
                
                {!isDate && field.choices && (
                    <div>
                        <p style={{ fontSize: '12px', color: '#605e5c', marginBottom: 12 }}>
                            Select a background color for each choice.
                        </p>
                        {field.choices.map(c => (
                            <div key={c} style={{ display: 'flex', alignItems: 'center', marginBottom: 10, gap: 15 }}>
                                <div style={{ width: 150, fontWeight: 500, color: '#323130' }}>{c}</div>
                                <input 
                                    type="color" 
                                    value={(this.state.choiceConfig && this.state.choiceConfig[c]) || '#ffffff'} 
                                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => this._saveChoiceColor(c, e.target.value)} 
                                    style={{ 
                                        border: '1px solid #c8c6c4', 
                                        borderRadius: '2px', 
                                        padding: '0', 
                                        width: '40px', 
                                        height: '30px', 
                                        cursor: 'pointer' 
                                    }}
                                />
                            </div>
                        ))}
                    </div>
                )}

                {isDate && (
                    <div>
                        <p style={{ fontSize: '12px', color: '#605e5c', marginBottom: 12 }}>
                            Define rules to highlight dates based on today's date.
                        </p>
                        {(this.state.dateRules || []).map((rule, idx) => (
                            <div key={idx} style={{ padding: 12, background: '#f3f2f1', marginBottom: 10, borderRadius: 4, border: '1px solid #e1dfdd' }}>
                                <Dropdown
                                    label="Condition"
                                    selectedKey={rule.condition}
                                    options={[
                                        { key: 'past', text: 'In the Past' },
                                        { key: 'future', text: 'In the Future' },
                                        { key: 'past_n', text: 'Older than N days' },
                                        { key: 'future_n', text: 'Due in N days' },
                                        { key: 'today', text: 'Today' }
                                    ]}
                                    onChange={(_, opt) => {
                                        const rules = [...this.state.dateRules!];
                                        rules[idx].condition = opt?.key as any;
                                        this.setState({ dateRules: rules });
                                    }}
                                    styles={{ root: { marginBottom: 8 } }}
                                />
                                {(rule.condition === 'past_n' || rule.condition === 'future_n') && (
                                    <TextField 
                                        label="Number of Days (N)" 
                                        type="number" 
                                        value={rule.days?.toString()} 
                                        onChange={(_, val) => {
                                            const rules = [...this.state.dateRules!];
                                            rules[idx].days = parseInt(val || '0', 10);
                                            this.setState({ dateRules: rules });
                                        }} 
                                    />
                                )}
                                <div style={{ marginTop: 12, display: 'flex', alignItems: 'center', gap: 10 }}>
                                    <label style={{ fontWeight: 600, fontSize: '13px' }}>Highlight Color:</label>
                                    <input 
                                        type="color" 
                                        value={rule.color} 
                                        onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                                            const rules = [...this.state.dateRules!];
                                            rules[idx].color = e.target.value;
                                            this.setState({ dateRules: rules });
                                        }} 
                                        style={{ border: '1px solid #c8c6c4', borderRadius: '2px', width: '30px', height: '30px', cursor: 'pointer', padding: 0 }}
                                    />
                                    <IconButton 
                                        iconProps={{ iconName: 'Delete' }} 
                                        title="Remove Rule" 
                                        ariaLabel="Remove Rule"
                                        onClick={() => {
                                            const rules = this.state.dateRules!.filter((_, i) => i !== idx);
                                            this.setState({ dateRules: rules });
                                        }} 
                                        styles={{ root: { color: '#d13438', marginLeft: 'auto' }, rootHovered: { background: '#fde7e9' } }}
                                    />
                                </div>
                            </div>
                        ))}
                        <PrimaryButton 
                            text="+ Add Date Rule" 
                            onClick={this._addDateRule} 
                            style={{ marginTop: 10, background: '#ffffff', color: '#0078d4', border: '1px solid #0078d4' }} 
                        />
                    </div>
                )}

                <div style={{ marginTop: 24, display: 'flex', gap: 10, borderTop: '1px solid #edebe9', paddingTop: 16 }}>
                    <PrimaryButton text="Save Formatting" onClick={() => this.props.onSave(this.state)} />
                    <DefaultButton text="Cancel" onClick={this.props.onCancel} />
                </div>
            </div>
        );
    }
}
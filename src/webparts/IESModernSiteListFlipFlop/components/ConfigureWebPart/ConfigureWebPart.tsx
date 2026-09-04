import * as React from 'react';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import styles from './ConfigureWebPart.module.scss';



export interface IConfigureWebPartProps {
    webPartContext: IWebPartContext;
    title: string;
    description?: string;
    buttonText?: string;
}


const ConfigureWebPart: React.SFC<IConfigureWebPartProps> = (props) => {

    const {
        webPartContext,
        title,
        description,
        buttonText,
    } = props;

    const setupTExtLabel = 'Please click on \'Set Up\' button to configure Flip Card App.';

    return (
        <div>
            <div>{title}</div>
            <div className={styles.SetupText}>
                <MessageBar messageBarType={MessageBarType.info} >
                    {
                    description ? description : <label className={styles.SetupText}>{setupTExtLabel}</label>
                    }
                </MessageBar>
            </div>
            <div className={styles.button}>
                <PrimaryButton iconProps={{ iconName: 'Edit' }} onClick={(e) => { e.preventDefault(); webPartContext.propertyPane.open(); }}>
                    {buttonText ? buttonText : 'Set Up'}
                </PrimaryButton>
            </div>
        </div>
    );
};

export default ConfigureWebPart;
import * as React from 'react';
import styles from './SpfxParentChildLists.module.scss';
import type { ISpfxParentChildListsProps } from './ISpfxParentChildListsProps';
import { Modal, PrimaryButton } from '@fluentui/react';

interface ISpfxParentChildListsState {
  isModalOpen: boolean;
}

export default class SpfxParentChildLists extends React.Component<ISpfxParentChildListsProps, ISpfxParentChildListsState> {
  constructor(props: ISpfxParentChildListsProps) {
    super(props);
    this.state = {
      isModalOpen: false
    };
  }

  private showModal = (): void => {
    this.setState({ isModalOpen: true });
  }

  private hideModal = (): void => {
    this.setState({ isModalOpen: false });
  }

  public render(): React.ReactElement<ISpfxParentChildListsProps> {
    const { hasTeamsContext } = this.props;

    return (
      <section className={`${styles.spfxParentChildLists} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.container}>
          <h1>Parent-Child List CRUD Operations</h1>
          
          <PrimaryButton 
            text="Open Modal" 
            onClick={this.showModal} 
            className={styles.button}
          />

          <Modal
            isOpen={this.state.isModalOpen}
            onDismiss={this.hideModal}
            isBlocking={false}
            isDarkOverlay={true}
          >
            <div className={styles.modal}>
              <h2>Add New Item</h2>
              <p>Form will go here...</p>
              <PrimaryButton text="Close" onClick={this.hideModal} />
            </div>
          </Modal>
        </div>
      </section>
    );
  }
}


